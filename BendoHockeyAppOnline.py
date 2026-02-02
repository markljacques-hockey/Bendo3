import streamlit as st
import pandas as pd
import urllib.parse
import io

# --- 1. SETUP & CONFIG ---
st.set_page_config(page_title="Hockey Team Manager", layout="wide")

# --- 2. HELPER FUNCTIONS ---
def snake_draft(players):
    """Distributes players 1-by-1 in a Snake pattern (A, B, B, A)."""
    team_a, team_b = [], []
    players = players.reset_index(drop=True)
    
    for i, (_, player) in enumerate(players.iterrows()):
        if i % 4 == 0 or i % 4 == 3:
            team_a.append(player)
        else:
            team_b.append(player)
            
    # Return DataFrames with columns preserved
    cols = players.columns
    return pd.DataFrame(team_a, columns=cols), pd.DataFrame(team_b, columns=cols)

def format_team_list(df, team_name):
    if df.empty: return f"{team_name} (0 players):\n"
    txt = f"{team_name} ({len(df)} players):\n"
    # Sort: Position first, then Name
    if 'Position' in df.columns and 'Full Name' in df.columns:
        df_sorted = df.sort_values(by=['Position', 'Full Name'])
        for _, row in df_sorted.iterrows():
            txt += f"- {row['Full Name']} ({row['Position']})\n"
    return txt

def get_top_n_score(df, n):
    if df.empty or n <= 0: return 0
    return df.sort_values(by='Score', ascending=False).head(n)['Score'].sum()

def enforce_rivalries(df_a, df_b):
    """Ensures specific pairs are not on the same team."""
    pairs = [
        ("Mike Tonietto", "Jamie Devin"),
        ("Mark Hicks", "Gary Fera")
    ]
    
    logs = []

    for p1, p2 in pairs:
        # Check if players exist in the current rosters (case-insensitive)
        # We create temporary lower-case columns for matching to be safe
        df_a['Name_Lower'] = df_a['Full Name'].astype(str).str.lower().str.strip()
        df_b['Name_Lower'] = df_b['Full Name'].astype(str).str.lower().str.strip()
        
        p1_key = p1.lower().strip()
        p2_key = p2.lower().strip()
        
        # Locate players
        p1_in_a = not df_a[df_a['Name_Lower'] == p1_key].empty
        p1_in_b = not df_b[df_b['Name_Lower'] == p1_key].empty
        p2_in_a = not df_a[df_a['Name_Lower'] == p2_key].empty
        p2_in_b = not df_b[df_b['Name_Lower'] == p2_key].empty
        
        # Only proceed if BOTH are playing
        if (p1_in_a or p1_in_b) and (p2_in_a or p2_in_b):
            collision = False
            team_with_both = None
            other_team = None
            
            if p1_in_a and p2_in_a:
                collision = True
                team_with_both = df_a
                other_team = df_b
            elif p1_in_b and p2_in_b:
                collision = True
                team_with_both = df_b
                other_team = df_a
            
            if collision:
                # We move P2 to the other team
                # 1. Get P2's details (Position)
                p2_row = team_with_both[team_with_both['Name_Lower'] == p2_key].iloc[0]
                p2_pos = p2_row['Position']
                
                # 2. Find a swap candidate in Other Team with SAME Position
                # We try to avoid swapping out the *other* constraint players
                protected = [x.lower() for x in [p1, p2] + [n for pair in pairs for n in pair]]
                
                candidates = other_team[
                    (other_team['Position'] == p2_pos) & 
                    (~other_team['Name_Lower'].isin(protected))
                ]
                
                if candidates.empty:
                    # If no safe candidate, take anyone of that position
                    candidates = other_team[other_team['Position'] == p2_pos]
                
                if not candidates.empty:
                    # Take the last one (lowest ranked usually) to swap
                    swap_target = candidates.iloc[-1]
                    
                    # 3. Perform Swap
                    # Remove P2 from Team With Both
                    team_with_both = team_with_both[team_with_both['Name_Lower'] != p2_key]
                    # Remove Target from Other Team
                    other_team = other_team[other_team['Name_Lower'] != swap_target['Name_Lower']]
                    
                    # Add P2 to Other Team
                    other_team = pd.concat([other_team, p2_row.to_frame().T], ignore_index=True)
                    # Add Target to Team With Both
                    team_with_both = pd.concat([team_with_both, swap_target.to_frame().T], ignore_index=True)
                    
                    logs.append(f"Separated **{p1} & {p2}**: Swapped {p2} with {swap_target['Full Name']}.")
                    
                    # Assign back to main dataframes
                    if p1_in_a and p2_in_a:
                        df_a, df_b = team_with_both, other_team
                    else:
                        df_b, df_a = team_with_both, other_team

        # Clean up temp columns
        if 'Name_Lower' in df_a.columns: del df_a['Name_Lower']
        if 'Name_Lower' in df_b.columns: del df_b['Name_Lower']

    return df_a, df_b, logs

# --- 3. MAIN APP ---
st.title("🏒 Hockey Team Manager & Generator")
st.markdown("""
**Instructions:**
1. **Upload** your Master Excel List.
2. **Select** the Sheet (e.g., Monday or Wednesday).
3. **Edit the List Below:** Change Availability to 'Yes', add new players, or update scores directly in the grid.
4. **Generate Teams:** The app uses the data exactly as it appears in the grid.
5. **Save Changes:** Download the updated Excel file at the bottom to keep your changes for next week.
""")

# --- FILE UPLOADER ---
uploaded_file = st.file_uploader("Upload Master Excel File", type=['xlsx', 'xlsm', 'xls'])

if uploaded_file is not None:
    try:
        # Load Excel File
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        # Select Sheet
        selected_sheet = st.selectbox("Select Roster (Sheet):", sheet_names)
        
        # Read Data
        raw_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=1).fillna("")

        # --- INTERACTIVE DATA EDITOR ---
        st.subheader(f"📝 Manage Roster: {selected_sheet}")
        st.info("💡 Tip: You can type directly in these cells. Click the **+** icon to add a new player.")
        
        edited_df = st.data_editor(
            raw_df,
            num_rows="dynamic",
            use_container_width=True,
            height=400,
            key="editor"
        )
        
        # --- PREPARE DATA ---
        req_cols = ['Availability', 'Reg/Spare', 'First_name', 'Last_name', '1st Choice', 'Score']
        missing = [c for c in req_cols if c not in edited_df.columns]
        if missing:
            st.error(f"Error: Missing columns: {', '.join(missing)}")
            st.stop()

        proc_df = edited_df.copy()
        proc_df['Availability'] = proc_df['Availability'].astype(str).str.strip().str.title()
        proc_df['1st Choice'] = proc_df['1st Choice'].astype(str).str.strip().str.upper()
        proc_df['Score'] = pd.to_numeric(proc_df['Score'], errors='coerce').fillna(0)
        proc_df['Full Name'] = proc_df['First_name'].astype(str).str.strip() + ' ' + proc_df['Last_name'].astype(str).str.strip()
        
        if '2nd Choice' in proc_df.columns:
            proc_df['2nd Choice'] = proc_df['2nd Choice'].astype(str).str.strip().str.upper()
        else:
            proc_df['2nd Choice'] = ''
            
        if 'Email' in proc_df.columns:
            proc_df['Email'] = proc_df['Email'].astype(str).str.strip()
        else:
            proc_df['Email'] = ''

        # --- GENERATE TEAMS BUTTON ---
        st.divider()
        col_gen, col_download = st.columns([1, 1])
        
        generate_clicked = False
        with col_gen:
            if st.button("🚀 Generate Teams using Table Above", type="primary"):
                generate_clicked = True

        # --- LOGIC ---
        if generate_clicked:
            available = proc_df[proc_df['Availability'] == 'Yes'].copy()
            
            if available.empty:
                st.error("No players marked as 'Yes' in the table above.")
                st.stop()

            # Dynamic Targets
            total_players = len(available)
            BASE_F, BASE_D, MIN_D_CRITICAL = 12, 8, 6

            if total_players <= 20:
                target_f, target_d = BASE_F, BASE_D
            else:
                extras = total_players - 20
                add_to_f = min(extras, 6)
                extras -= add_to_f
                add_to_d = min(extras, 4)
                target_f, target_d = BASE_F + add_to_f, BASE_D + add_to_d
            
            st.success(f"**Strategy:** {total_players} Players -> Aiming for {target_f} F / {target_d} D")

            # Sort
            available = available.sample(frac=1).reset_index(drop=True)
            available['Status_Rank'] = available['Reg/Spare'].apply(lambda x: 0 if str(x).strip().upper() == 'R' else 1)
            available = available.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])

            pool_d = available[available['1st Choice'] == 'D'].copy()
            pool_f = available[available['1st Choice'] == 'F'].copy()

            # Fill Gaps
            if len(pool_d) < MIN_D_CRITICAL:
                needed = MIN_D_CRITICAL - len(pool_d)
                cands = pool_f[pool_f['2nd Choice'] == 'D']
                if not cands.empty:
                    conv = cands.head(needed)
                    pool_d = pd.concat([pool_d, conv])
                    pool_f = pool_f.drop(conv.index)
                    st.warning(f"⚠️ Moved F -> D (Critical): {', '.join(conv['Full Name'])}")

            if len(pool_f) < target_f:
                needed = target_f - len(pool_f)
                surplus_d = len(pool_d) - MIN_D_CRITICAL
                if surplus_d > 0:
                    conv = pool_d[pool_d['2nd Choice'] == 'F'].head(min(needed, surplus_d))
                    if not conv.empty:
                        pool_f = pd.concat([pool_f, conv])
                        pool_d = pool_d.drop(conv.index)
                        st.info(f"Moved D -> F: {', '.join(conv['Full Name'])}")

            if len(pool_d) < target_d:
                needed = target_d - len(pool_d)
                surplus_f = len(pool_f) - target_f
                if surplus_f > 0:
                    conv = pool_f[pool_f['2nd Choice'] == 'D'].head(min(needed, surplus_f))
                    if not conv.empty:
                        pool_d = pd.concat([pool_d, conv])
                        pool_f = pool_f.drop(conv.index)
                        st.info(f"Moved F -> D: {', '.join(conv['Full Name'])}")

            # Draft
            pool_d = pool_d.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])
            pool_f = pool_f.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])
            
            sel_d = pool_d.head(target_d).copy()
            sel_f = pool_f.head(target_f).copy()
            cuts_d = pool_d.iloc[target_d:].copy()
            cuts_f = pool_f.iloc[target_f:].copy()
            
            sel_d['Position'] = 'D'
            sel_f['Position'] = 'F'
            
            da, db = snake_draft(sel_d)
            fa, fb = snake_draft(sel_f)
            
            team_a = pd.concat([da, fa], ignore_index=True).sample(frac=1).reset_index(drop=True)
            team_b = pd.concat([db, fb], ignore_index=True).sample(frac=1).reset_index(drop=True)

            # --- NEW STEP: ENFORCE RIVALRIES ---
            team_a, team_b, rivalry_logs = enforce_rivalries(team_a, team_b)
            
            if rivalry_logs:
                st.divider()
                for log in rivalry_logs:
                    st.success(f"⚖️ {log}")

            # Display
            score_a = team_a['Score'].sum() if not team_a.empty else 0
            score_b = team_b['Score'].sum() if not team_b.empty else 0
            common = min(len(team_a), len(team_b))
            fair_a = get_top_n_score(team_a, common)
            fair_b = get_top_n_score(team_b, common)

            c1, c2 = st.columns(2)
            disp_cols = ['Full Name', 'Position']
            
            with c1:
                st.header("🔴 Red Team")
                d_cnt = len(team_a[team_a['Position']=='D'])
                f_cnt = len(team_a[team_a['Position']=='F'])
                st.write(f"**Score:** {score_a} (Top {common}: {fair_a})")
                st.write(f"Players: {len(team_a)} ({d_cnt} D / {f_cnt} F)")
                st.dataframe(team_a[disp_cols], hide_index=True)
                
            with c2:
                st.header("⚪ White Team")
                d_cnt = len(team_b[team_b['Position']=='D'])
                f_cnt = len(team_b[team_b['Position']=='F'])
                st.write(f"**Score:** {score_b} (Top {common}: {fair_b})")
                st.write(f"Players: {len(team_b)} ({d_cnt} D / {f_cnt} F)")
                st.dataframe(team_b[disp_cols], hide_index=True)

            if not cuts_d.empty or not cuts_f.empty:
                st.error("🚫 Cuts/Spares:")
                for _, r in cuts_d.iterrows(): st.write(f"- {r['Full Name']} (D)")
                for _, r in cuts_f.iterrows(): st.write(f"- {r['Full Name']} (F)")

            # Email
            all_p = pd.concat([team_a, team_b])
            recipients = [x for x in all_p['Email'].unique() if x and str(x).strip()]
            bcc = ",".join(recipients)
            body = f"Hello everyone,\n\nHere are the rosters:\n\n{format_team_list(team_a, 'RED TEAM')}\n{format_team_list(team_b, 'WHITE TEAM')}\nKeep your sticks on the ice!"
            
            subj = f"Bendo Hockey Lineups - {selected_sheet}"
            link_subj = urllib.parse.quote(subj)
            link_body = urllib.parse.quote(body)
            link_bcc = urllib.parse.quote(bcc)
            
            st.subheader("📧 Email Tools")
            st.text_area("Copy Text:", value=body, height=150)
            
            c_email1, c_email2 = st.columns(2)
            with c_email1:
                st.link_button("🚀 Open Default Mail App", f"mailto:?bcc={link_bcc}&subject={link_subj}&body={link_body}")
            with c_email2:
                st.link_button("📧 Draft in Gmail (Web)", f"https://mail.google.com/mail/?view=cm&fs=1&bcc={link_bcc}&su={link_subj}&body={link_body}")

        # --- DOWNLOAD ---
        with col_download:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                edited_df.to_excel(writer, index=False, sheet_name=selected_sheet, startrow=1)
            
            st.download_button(
                label=f"💾 Download Updated '{selected_sheet}' List",
                data=output.getvalue(),
                file_name=f"Updated_{selected_sheet}_List.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")