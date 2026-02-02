import streamlit as st
import pandas as pd
import urllib.parse
import io

# --- 1. SETUP & HELPER FUNCTIONS ---
st.set_page_config(page_title="Hockey Team Balancer")

def snake_draft(players):
    """Distributes players 1-by-1 in a Snake pattern (A, B, B, A)."""
    team_a, team_b = [], []
    players = players.reset_index(drop=True)
    
    for i, (_, player) in enumerate(players.iterrows()):
        if i % 4 == 0 or i % 4 == 3:
            team_a.append(player)
        else:
            team_b.append(player)
            
    # Return DataFrames
    cols = players.columns
    if not team_a: df_a = pd.DataFrame(columns=cols)
    else: df_a = pd.DataFrame(team_a, columns=cols)
        
    if not team_b: df_b = pd.DataFrame(columns=cols)
    else: df_b = pd.DataFrame(team_b, columns=cols)
        
    return df_a, df_b

def format_team_list(df, team_name):
    if df.empty: return f"{team_name} (0 players):\n"
    txt = f"{team_name} ({len(df)} players):\n"
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
        # Create temp lower-case columns for matching
        df_a['Name_Lower'] = df_a['Full Name'].astype(str).str.lower().str.strip()
        df_b['Name_Lower'] = df_b['Full Name'].astype(str).str.lower().str.strip()
        
        p1_key = p1.lower().strip()
        p2_key = p2.lower().strip()
        
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
                # Find P2 and Swap
                p2_row = team_with_both[team_with_both['Name_Lower'] == p2_key].iloc[0]
                p2_pos = p2_row['Position']
                
                # Find candidate in other team (same pos, avoid other rivals)
                protected = [x.lower() for x in [p1, p2] + [n for pair in pairs for n in pair]]
                candidates = other_team[
                    (other_team['Position'] == p2_pos) & 
                    (~other_team['Name_Lower'].isin(protected))
                ]
                
                if candidates.empty:
                    candidates = other_team[other_team['Position'] == p2_pos]
                
                if not candidates.empty:
                    swap_target = candidates.iloc[-1] # Swap lowest ranked if possible
                    
                    # Remove
                    team_with_both = team_with_both[team_with_both['Name_Lower'] != p2_key]
                    other_team = other_team[other_team['Name_Lower'] != swap_target['Name_Lower']]
                    
                    # Add (Convert Series to DataFrame before concat to avoid warnings)
                    other_team = pd.concat([other_team, p2_row.to_frame().T], ignore_index=True)
                    team_with_both = pd.concat([team_with_both, swap_target.to_frame().T], ignore_index=True)
                    
                    logs.append(f"Separated **{p1} & {p2}**: Swapped {p2} with {swap_target['Full Name']}.")
                    
                    if p1_in_a and p2_in_a:
                        df_a, df_b = team_with_both, other_team
                    else:
                        df_b, df_a = team_with_both, other_team

    # Clean up
    if 'Name_Lower' in df_a.columns: del df_a['Name_Lower']
    if 'Name_Lower' in df_b.columns: del df_b['Name_Lower']

    return df_a, df_b, logs

# --- 2. MAIN APP INTERFACE ---
st.title("🏒 Hockey Team Generator")
st.write("Upload your Excel player sheet.")

# File Uploader
uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls', 'xlsm'])

if uploaded_file is not None:
    try:
        # 1. Get Sheet Names
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        selected_sheet = st.selectbox("Select the Sheet to use:", sheet_names)
        
        # 2. Load Data
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=1)
        
        # Required Columns
        required_cols = ['Availability', 'Reg/Spare', 'First_name', 'Last_name', '1st Choice', 'Score']
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.error(f"Missing required columns in sheet '{selected_sheet}': {', '.join(missing)}")
            st.stop()

        # Clean Data
        df['Availability'] = df['Availability'].astype(str).str.strip().str.title()
        df['1st Choice'] = df['1st Choice'].astype(str).str.strip().str.upper()
        df['Full Name'] = df['First_name'].astype(str).str.strip() + ' ' + df['Last_name'].astype(str).str.strip()
        
        if '2nd Choice' in df.columns:
            df['2nd Choice'] = df['2nd Choice'].fillna('').astype(str).str.strip().str.upper()
        else:
            df['2nd Choice'] = ''
            
        if 'Email' not in df.columns:
            df['Email'] = ''
        else:
            df['Email'] = df['Email'].fillna('').astype(str).str.strip()

        # Filter Available
        available = df[df['Availability'] == 'Yes'].copy()
        
        if available.empty:
            st.error(f"No players marked as 'Yes' in sheet '{selected_sheet}'.")
            st.stop()
        
        # --- 3. DYNAMIC TARGET CALCULATION ---
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
        
        st.info(f"**Roster Strategy ({selected_sheet}):** Found {total_players} players. Aiming for **{target_f} Forwards** and **{target_d} Defensemen**.")

        # --- 4. RANDOMIZE & SORT ---
        available = available.sample(frac=1).reset_index(drop=True)
        available['Status_Rank'] = available['Reg/Spare'].apply(lambda x: 0 if str(x).strip().upper() == 'R' else 1)
        available = available.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])

        pool_d = available[available['1st Choice'] == 'D'].copy()
        pool_f = available[available['1st Choice'] == 'F'].copy()

        # --- 5. FILL GAPS ---
        # Critical D
        if len(pool_d) < MIN_D_CRITICAL:
            needed = MIN_D_CRITICAL - len(pool_d)
            candidates = pool_f[pool_f['2nd Choice'] == 'D']
            if not candidates.empty:
                converts = candidates.head(needed)
                pool_d = pd.concat([pool_d, converts])
                pool_f = pool_f.drop(converts.index)
                moved_names = ", ".join(converts['Full Name'].tolist())
                st.warning(f"⚠️ Critical D Shortage: Moved {len(converts)} player(s) from F to D: **{moved_names}**")
        
        # Fill Forwards
        if len(pool_f) < target_f:
            needed = target_f - len(pool_f)
            surplus_d = len(pool_d) - MIN_D_CRITICAL
            if surplus_d > 0:
                can_take = min(needed, surplus_d)
                candidates = pool_d[pool_d['2nd Choice'] == 'F']
                if not candidates.empty:
                    converts = candidates.head(can_take)
                    pool_f = pd.concat([pool_f, converts])
                    pool_d = pool_d.drop(converts.index)
                    moved_names = ", ".join(converts['Full Name'].tolist())
                    st.info(f"Moved {len(converts)} player(s) from D to F: **{moved_names}**")

        # Fill Defense
        if len(pool_d) < target_d:
            d_shortage = target_d - len(pool_d)
            surplus_f = len(pool_f) - target_f
            if surplus_f > 0:
                amount_to_move = min(d_shortage, surplus_f)
                candidates = pool_f[pool_f['2nd Choice'] == 'D']
                if not candidates.empty:
                    converts = candidates.head(amount_to_move)
                    pool_d = pd.concat([pool_d, converts])
                    pool_f = pool_f.drop(converts.index)
                    moved_names = ", ".join(converts['Full Name'].tolist())
                    st.info(f"Moved {len(converts)} player(s) from F to D: **{moved_names}**")

        # --- 6. LIMITS & CUTS ---
        pool_d = pool_d.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])
        pool_f = pool_f.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])

        selected_d = pool_d.head(target_d).copy()
        selected_f = pool_f.head(target_f).copy()
        cuts_d = pool_d.iloc[target_d:].copy()
        cuts_f = pool_f.iloc[target_f:].copy()

        selected_d['Position'] = 'D'
        selected_f['Position'] = 'F'

        # --- 7. DRAFT TEAMS ---
        d_a, d_b = snake_draft(selected_d)
        f_a, f_b = snake_draft(selected_f)

        team_a = pd.concat([d_a, f_a], ignore_index=True).sample(frac=1).reset_index(drop=True)
        team_b = pd.concat([d_b, f_b], ignore_index=True).sample(frac=1).reset_index(drop=True)

        # --- 8. ENFORCE RIVALRIES ---
        team_a, team_b, rivalry_logs = enforce_rivalries(team_a, team_b)

        # --- 9. DISPLAY ---
        if st.button("Shuffle Teams Again"):
            st.rerun()

        if rivalry_logs:
            st.divider()
            for log in rivalry_logs:
                st.success(f"⚖️ {log}")

        # Recalculate Scores after potential swaps
        count_a, count_b = len(team_a), len(team_b)
        common_count = min(count_a, count_b)
        total_score_a = team_a['Score'].sum() if not team_a.empty else 0
        total_score_b = team_b['Score'].sum() if not team_b.empty else 0
        fair_score_a = get_top_n_score(team_a, common_count)
        fair_score_b = get_top_n_score(team_b, common_count)

        cols = ['Full Name', 'Position']
        col1, col2 = st.columns(2)
        
        with col1:
            st.header(f"🔴 Red Team")
            cnt_d_a = len(team_a[team_a['Position'] == 'D'])
            cnt_f_a = len(team_a[team_a['Position'] == 'F'])
            st.write(f"**Total Score:** {total_score_a}")
            if common_count > 0: st.write(f"**Top {common_count} Score:** {fair_score_a}")
            st.write(f"Players: {len(team_a)} **({cnt_d_a} D / {cnt_f_a} F)**")
            if not team_a.empty: st.dataframe(team_a[cols], hide_index=True)
            else: st.write("No players.")
                
        with col2:
            st.header(f"⚪ White Team")
            cnt_d_b = len(team_b[team_b['Position'] == 'D'])
            cnt_f_b = len(team_b[team_b['Position'] == 'F'])
            st.write(f"**Total Score:** {total_score_b}")
            if common_count > 0: st.write(f"**Top {common_count} Score:** {fair_score_b}")
            st.write(f"Players: {len(team_b)} **({cnt_d_b} D / {cnt_f_b} F)**")
            if not team_b.empty: st.dataframe(team_b[cols], hide_index=True)
            else: st.write("No players.")

        # Cuts
        if not cuts_d.empty or not cuts_f.empty:
            st.divider()
            st.subheader("🚫 Undrafted Players")
            c1, c2 = st.columns(2)
            with c1:
                if not cuts_d.empty:
                    st.error(f"**Defense Cuts ({len(cuts_d)}):**")
                    for name in cuts_d['Full Name']: st.write(f"- {name}")
                else: st.success("No Defense cuts.")
            with c2:
                if not cuts_f.empty:
                    st.error(f"**Forward Cuts ({len(cuts_f)}):**")
                    for name in cuts_f['Full Name']: st.write(f"- {name}")
                else: st.success("No Forward cuts.")

        # --- 10. EMAIL ---
        st.divider()
        st.subheader("📧 Notify Players")
        all_players = pd.concat([team_a, team_b])
        if not all_players.empty:
            recipients = [e for e in all_players['Email'].unique() if e != '' and pd.notna(e)]
            bcc_string = ",".join(recipients)
            email_body = f"""Hello everyone,\n\nHere are the rosters for the upcoming game:\n\n{format_team_list(team_a, "RED TEAM")}\n{format_team_list(team_b, "WHITE TEAM")}\nKeep your sticks on the ice!"""
            st.text_area("Email Text (Draft Only):", value=email_body, height=300)

            subject_line = f"Bendo Hockey Lineups - {selected_sheet}"
            safe_subject = urllib.parse.quote(subject_line)
            safe_body = urllib.parse.quote(email_body)
            safe_bcc = urllib.parse.quote(bcc_string)
            mailto_url = f"mailto:?bcc={safe_bcc}&subject={safe_subject}&body={safe_body}"
            gmail_url = f"https://mail.google.com/mail/?view=cm&fs=1&bcc={safe_bcc}&su={safe_subject}&body={safe_body}"

            if len(recipients) > 0:
                col_email1, col_email2 = st.columns(2)
                with col_email1: st.link_button("🚀 Open Default Email App", mailto_url)
                with col_email2: st.link_button("📧 Draft in Gmail (Web)", gmail_url)
            else: st.caption("No emails found to generate link.")
                
    except Exception as e:
        st.error(f"Error processing file: {e}")