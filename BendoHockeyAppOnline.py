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
        # Pattern: A, B, B, A
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

        # --- 6. PRE-ASSIGN RIVALS (NEW LOGIC) ---
        # We assign them NOW, before the main draft, to ensure they are separated 
        # without disrupting the scoring balance of the rest of the draft.
        
        pre_team_a = []
        pre_team_b = []
        rivalry_notes = []

        # Helper to find and remove a player from the pools
        def extract_player(name, pd_pool, pf_pool):
            name_key = name.lower().strip()
            # Check D
            for idx, row in pd_pool.iterrows():
                if row['Full Name'].lower().strip() == name_key:
                    pd_pool = pd_pool.drop(idx)
                    row['Position'] = 'D'
                    return row, pd_pool, pf_pool
            # Check F
            for idx, row in pf_pool.iterrows():
                if row['Full Name'].lower().strip() == name_key:
                    pf_pool = pf_pool.drop(idx)
                    row['Position'] = 'F'
                    return row, pd_pool, pf_pool
            return None, pd_pool, pf_pool

        # Rival Pairs
        rival_pairs = [
            ("Mike Tonietto", "Jamie Devin"),
            ("Mark Hicks", "Gary Fera")
        ]

        # Use an index to alternate which team gets the 'better' player first to balance pre-seeds
        pair_index = 0 

        for p1_name, p2_name in rival_pairs:
            # Check existence in current pools
            # We assume names are unique enough. 
            # We do a 'dry run' check first to see if both exist
            
            p1_found = any(df['Full Name'].str.lower().str.strip() == p1_name.lower().strip() for idx, df in [pool_d, pool_f] for _, row in df.iterrows())
            p2_found = any(df['Full Name'].str.lower().str.strip() == p2_name.lower().strip() for idx, df in [pool_d, pool_f] for _, row in df.iterrows())
            
            # Note: The above boolean check is slightly inefficient but safe. 
            # Better to try extracting and putting back if not found? 
            # Actually, let's just try to extract both.
            
            if True: # We just attempt extraction
                p1_obj, pool_d, pool_f = extract_player(p1_name, pool_d, pool_f)
                p2_obj, pool_d, pool_f = extract_player(p2_name, pool_d, pool_f)

                if p1_obj is not None and p2_obj is not None:
                    # BOTH are playing. Force Separate.
                    
                    # Sort by score to distribute fairly
                    pair_objs = sorted([p1_obj, p2_obj], key=lambda x: x['Score'], reverse=True)
                    higher = pair_objs[0]
                    lower = pair_objs[1]
                    
                    # Alternate assignment to balance the "Seed" scores
                    if pair_index % 2 == 0:
                        pre_team_a.append(higher)
                        pre_team_b.append(lower)
                        rivalry_notes.append(f"Separated {p1_name} & {p2_name}: {higher['Full Name']} -> Red, {lower['Full Name']} -> White")
                    else:
                        pre_team_b.append(higher)
                        pre_team_a.append(lower)
                        rivalry_notes.append(f"Separated {p1_name} & {p2_name}: {higher['Full Name']} -> White, {lower['Full Name']} -> Red")
                    
                    pair_index += 1
                    
                else:
                    # One or both missing. Put them back if we found one!
                    if p1_obj is not None:
                        if p1_obj['Position'] == 'D': pool_d = pd.concat([pool_d, p1_obj.to_frame().T])
                        else: pool_f = pd.concat([pool_f, p1_obj.to_frame().T])
                    if p2_obj is not None:
                        if p2_obj['Position'] == 'D': pool_d = pd.concat([pool_d, p2_obj.to_frame().T])
                        else: pool_f = pd.concat([pool_f, p2_obj.to_frame().T])

        # --- 7. DRAFT THE REMAINDER ---
        # Sort remaining pools again
        pool_d = pool_d.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])
        pool_f = pool_f.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])

        # Determine how many slots are left to fill
        # We subtract the pre-assigned players from the targets
        pre_d_a = len([p for p in pre_team_a if p['Position'] == 'D'])
        pre_f_a = len([p for p in pre_team_a if p['Position'] == 'F'])
        
        # We just grab everyone remaining. The snake draft handles the count.
        # But we need to assign positions for display before drafting
        selected_d = pool_d.head(target_d).copy() # Note: This logic caps total D. 
        # If we already have D in pre-teams, we need to be careful not to over-draft.
        # Actually, 'pool_d' only contains UNASSIGNED players now.
        # So we should take (Total_Target - Assigned) from pool.
        
        total_pre_d = len([p for p in pre_team_a + pre_team_b if p['Position'] == 'D'])
        total_pre_f = len([p for p in pre_team_a + pre_team_b if p['Position'] == 'F'])
        
        needed_d = max(0, target_d - total_pre_d)
        needed_f = max(0, target_f - total_pre_f)
        
        selected_d = pool_d.head(needed_d).copy()
        selected_f = pool_f.head(needed_f).copy()
        
        # Identify Cuts
        cuts_d = pool_d.iloc[needed_d:].copy()
        cuts_f = pool_f.iloc[needed_f:].copy()

        selected_d['Position'] = 'D'
        selected_f['Position'] = 'F'

        # Run Snake Draft on Remainder
        d_a, d_b = snake_draft(selected_d)
        f_a, f_b = snake_draft(selected_f)

        # --- 8. COMBINE PRE-ASSIGNED + DRAFTED ---
        # Convert pre-lists to dataframes
        cols = df.columns # Use original columns
        
        if pre_team_a: df_pre_a = pd.DataFrame(pre_team_a, columns=cols)
        else: df_pre_a = pd.DataFrame(columns=cols)
            
        if pre_team_b: df_pre_b = pd.DataFrame(pre_team_b, columns=cols)
        else: df_pre_b = pd.DataFrame(columns=cols)

        # Merge
        team_a = pd.concat([df_pre_a, d_a, f_a], ignore_index=True)
        team_b = pd.concat([df_pre_b, d_b, f_b], ignore_index=True)
        
        # Shuffle final result so the "Pre-assigned" players aren't always at the top of the list
        team_a = team_a.sample(frac=1).reset_index(drop=True)
        team_b = team_b.sample(frac=1).reset_index(drop=True)

        # --- 9. DISPLAY ---
        if st.button("Shuffle Teams Again"):
            st.rerun()

        if rivalry_notes:
            st.divider()
            for note in rivalry_notes:
                st.success(f"⚖️ {note}")

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