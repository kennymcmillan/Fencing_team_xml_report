import streamlit as st
import xml.etree.ElementTree as ET
from collections import defaultdict
import pandas as pd
from io import BytesIO

# Set the page configuration
st.set_page_config(layout="wide")

# Function to extract team details and build a fencer dictionary
# Function to extract team details and build a fencer dictionary
def build_fencer_dict_and_team_dict(root):
    fencer_dict = {}
    team_dict = {}
    for equipe in root.findall(".//Equipe"):
        team_id = equipe.get('ID')
        nation = equipe.get('Nation')
        
        # Determine the team name based on the format of ID, handle None safely
        if team_id and team_id.isdigit():
            team_name = equipe.get('IdOrigine')
        else:
            team_name = team_id

        # Skip entries with missing or unknown values
        if not team_id or not nation or not team_name:
            continue

        # Build the team dictionary
        team_dict[team_id] = {
            "Team Name": team_name,
            "Nation": nation
        }

        # Build the fencer dictionary
        for tireur in equipe.findall(".//Tireur"):
            fencer_id = tireur.get('ID')
            fencer_name = f"{tireur.get('Prenom')} {tireur.get('Nom')}"
            date_of_birth = tireur.get("DateNaissance", "Unknown")
            lateralite = tireur.get("Lateralite", "Unknown")

            fencer_dict[fencer_id] = {
                "Name": fencer_name,
                "Date of Birth": date_of_birth,
                "Lateralite": lateralite,
                "EquipeID": team_id,
                "Nation": nation,
                "Team Name": team_name
            }
    return fencer_dict, team_dict

# Function to parse rankings and create DataFrame
def extract_final_rankings(root, team_dict):
    rankings = []
    for phase in root.findall(".//PhaseDeTableaux"):
        for equipe in phase.findall(".//Equipe"):
            team_id = equipe.get('REF')
            final_rank = equipe.get('RangFinal')
            
            # Only proceed if both team_id and final_rank are available
            if team_id and final_rank:
                # Get the team name from team_dict, default to the ID if not found
                team_name = team_dict.get(team_id, {}).get("Team Name", team_id)
                nation = team_dict.get(team_id, {}).get("Nation", "Unknown")

                rankings.append({
                    "Team Name": team_name,
                    "Nation": nation,
                    "Final Rank": int(final_rank)  # Convert rank to integer for sorting
                })

    # Convert to DataFrame and sort by Final Rank
    rankings_df = pd.DataFrame(rankings).sort_values(by="Final Rank").reset_index(drop=True)
    return rankings_df


# Sidebar file uploader
st.sidebar.header("Upload XML File")
uploaded_file = st.sidebar.file_uploader("Choose an XML file", type="xml")

if uploaded_file:
    # Parse the XML file
    tree = ET.parse(uploaded_file)
    root = tree.getroot()

    # Build the fencer and team dictionaries
    fencer_dict, team_dict = build_fencer_dict_and_team_dict(root)

    final_rankings_df = extract_final_rankings(root, team_dict)

    # Tabs for different sections
    tabs = st.tabs(["Overview", "Nation Overview", "Tables", "Review", "Final Rankings"])

    with tabs[0]:
        # Overview Tab
        st.header("Competition Overview")

        # Extract competition details
        annee = root.get("Annee", "Unknown Year")
        titre_court = root.get("TitreCourtTournoi") or root.get("TitreCourt") or "Unknown Tournament"
        championnat = root.get("Championnat", "Unknown Championship")
        categorie = root.get("Categorie", "Unknown Category")
        arme = root.get("Arme", "Unknown Weapon")
        sexe = root.get("Sexe", "Unknown Gender")
        date = root.get("Date", "Unknown Date")
        lieu = root.get("Lieu", "Unknown Location")

        # Map abbreviations to full names
        categorie_full = "Cadet" if categorie == "C" else "Junior" if categorie == "J" else categorie
        arme_full = "Epee" if arme == "E" else "Sabre" if arme == "S" else "Foil" if arme == "F" else arme
        sexe_full = "Male" if sexe == "M" else "Female" if sexe == "F" else sexe

        # Display competition header
        st.header(f"{annee} - {titre_court} - {championnat}")
        st.subheader(f"Date: {date}, Location: {lieu}")
        st.subheader(f"Category: {categorie_full}, Weapon: {arme_full}, Gender: {sexe_full}")
        st.subheader("-----------------------------------------------------------------------------------------------------")

        # Display final ranking of Qatar team if available
        qatar_ranking = final_rankings_df[final_rankings_df['Nation'] == "QAT"]  # Filter for Qatar
        if not qatar_ranking.empty:
            qatar_final_rank = qatar_ranking.iloc[0]['Final Rank']  # Get the first row's rank
            st.subheader(f"The Qatar team achieved a final ranking of {qatar_final_rank}.")
        else:
            st.subheader("No final ranking data available for the Qatar team.")

        st.subheader("-----------------------------------------------------------------------------------------------------")

        # Display Qatari fencers
        qatari_fencers = [fencer for fencer in fencer_dict.values() if fencer["Nation"] == "QAT"]
        if qatari_fencers:
            st.markdown("### Qatari Fencers Participating")
            st.table(pd.DataFrame(qatari_fencers))
            st.subheader("-----------------------------------------------------------------------------------------------------")
        else:
            st.write("No Qatari fencers found in this competition.")
            st.subheader("-----------------------------------------------------------------------------------------------------")

        # Extract valid teams and countries
        teams = {team_id for team_id in team_dict}
        countries = {details['Nation'] for details in team_dict.values()}

        st.subheader(f"**Number of Teams:** {len(teams)}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;**Number of Countries:** {len(countries)}")
        st.subheader("-----------------------------------------------------------------------------------------------------")

        # Table of number of teams per country
        country_team_count = defaultdict(int)
        for details in team_dict.values():
            nation = details['Nation']
            if nation and len(nation) == 3:
                country_team_count[nation] += 1

        team_count_df = pd.DataFrame(list(country_team_count.items()), columns=["Country", "Number of Teams"])

        # Display table in 4 columns
        st.markdown("### Number of Teams per Country")
        num_columns = 4
        rows = (len(team_count_df) + num_columns - 1) // num_columns

        for row_idx in range(rows):
            cols = st.columns(num_columns)
            for col_idx, col in enumerate(cols):
                if row_idx + col_idx * rows < len(team_count_df):
                    country, count = team_count_df.iloc[row_idx + col_idx * rows]
                    col.metric(country, count)

    with tabs[1]:
        # Nation Overview Tab
        st.header("Nation Overview")

        # Extract all nations from the team_dict
        nations = sorted({details['Nation'] for details in team_dict.values()})
        selected_nation = st.selectbox("Select a Nation", nations)

        if selected_nation:
            # Filter fencers by selected nation using fencer_dict
            fencer_data = [fencer for fencer in fencer_dict.values() if fencer['Nation'] == selected_nation]

            # Convert to DataFrame and display
            if fencer_data:
                fencer_df = pd.DataFrame(fencer_data)
                st.table(fencer_df)
            else:
                st.write(f"No fencers found for {selected_nation}.")

    with tabs[2]:
        # Tables Tab
        st.header("Tables Overview")

        # Data structure to hold match results grouped by stage
        results_by_stage = defaultdict(list)
        export_data = []

        # Dictionary to accumulate touches and results for Qatari fencers
        accumulated_touches = defaultdict(lambda: {"scored": 0, "against": 0, "matches": []})

        # Iterate over all SuiteDeTableaux and Tableaux
        for suite in root.findall(".//SuiteDeTableaux"):
            for tableau in suite.findall(".//Tableau"):
                stage_title = tableau.get('Titre', 'Unknown Stage')

                # Reset touch_scores for each stage
                touch_scores = defaultdict(lambda: {"scored": 0, "against": 0})
                qatar_in_stage = False

                for match in tableau.findall(".//Match"):
                    team_d_ref = match.find(".//Equipe[@Cote='D']")
                    team_g_ref = match.find(".//Equipe[@Cote='G']")

                    # Resolve teams using team_dict
                    team_d = team_dict.get(team_d_ref.get('REF')) if team_d_ref is not None else None
                    team_g = team_dict.get(team_g_ref.get('REF')) if team_g_ref is not None else None

                    # Get team names
                    team_d_name = team_d["Team Name"] if team_d else "Unknown Team"
                    team_g_name = team_g["Team Name"] if team_g else "Unknown Team"

                    if (team_d and team_d["Nation"] == "QAT") or (team_g and team_g["Nation"] == "QAT"):
                        qatar_in_stage = True
                        prev_score_1, prev_score_2 = 0, 0

                        for assaut in match.findall(".//Assaut"):
                            fencer_d_ref = assaut.find(".//Tireur[@Cote='D']").get('REF')
                            fencer_g_ref = assaut.find(".//Tireur[@Cote='G']").get('REF')
                            score_d = int(assaut.find(".//Tireur[@Cote='D']").get('Score', 0))
                            score_g = int(assaut.find(".//Tireur[@Cote='G']").get('Score', 0))

                            # Calculate touches for this stage
                            touches_d = score_d - prev_score_1
                            touches_g = score_g - prev_score_2
                            diff_in_touches = touches_d - touches_g

                            prev_score_1, prev_score_2 = score_d, score_g

                            fencer_d = fencer_dict.get(fencer_d_ref, {"Name": "Unknown", "Nation": "Unknown"})
                            fencer_g = fencer_dict.get(fencer_g_ref, {"Name": "Unknown", "Nation": "Unknown"})

                            # Record touches and outcomes for Qatari fencers
                            if fencer_d["Nation"] == "QAT":
                                touch_scores[fencer_d["Name"]]["scored"] += touches_d
                                touch_scores[fencer_d["Name"]]["against"] += touches_g
                                accumulated_touches[fencer_d["Name"]]["scored"] += touches_d
                                accumulated_touches[fencer_d["Name"]]["against"] += touches_g
                                outcome = "Victory" if touches_d > touches_g else "Defeat" if touches_d < touches_g else "Draw"
                                accumulated_touches[fencer_d["Name"]]["matches"].append({
                                    "Opponent Team": team_g_name,
                                    "Outcome": outcome
                                })

                            if fencer_g["Nation"] == "QAT":
                                touch_scores[fencer_g["Name"]]["scored"] += touches_g
                                touch_scores[fencer_g["Name"]]["against"] += touches_d
                                accumulated_touches[fencer_g["Name"]]["scored"] += touches_g
                                accumulated_touches[fencer_g["Name"]]["against"] += touches_d
                                outcome = "Victory" if touches_g > touches_d else "Defeat" if touches_g < touches_d else "Draw"
                                accumulated_touches[fencer_g["Name"]]["matches"].append({
                                    "Opponent Team": team_d_name,
                                    "Outcome": outcome
                                })

                            row_data = {
                                "Team_1": team_d_name,
                                "Fencer_1": fencer_d["Name"],
                                "Touches_1": f"{touches_d} ({diff_in_touches})",
                                "Score_1": score_d,
                                "Score_2": score_g,
                                "Touches_2": f"{touches_g} ({-diff_in_touches})",
                                "Fencer_2": fencer_g["Name"],
                                "Team_2": team_g_name
                            }
                            results_by_stage[stage_title].append(row_data)
                            export_data.append([stage_title] + list(row_data.values()))

        # Display results by stage
        for stage_title, matches in results_by_stage.items():
            st.subheader(f"Stage: {stage_title}")
            st.table(matches)



    with tabs[3]:
        # Review Tab
        st.header("Review - Accumulated Touches and Match Outcomes for Qatari Fencers")

        qatari_summary = []
        for fencer, scores in accumulated_touches.items():
            total = scores["scored"] - scores["against"]
            qatari_summary.append({
                "Fencer": fencer,
                "Scored": scores["scored"],
                "Conceded": scores["against"],
                "Total": total,
                "Matches": scores["matches"]
            })

        # Display overall accumulated touches and aggregated match outcomes
        if qatari_summary:
            for fencer_data in qatari_summary:
                st.markdown(f"### Fencer: {fencer_data['Fencer']}")
                st.write(f"Scored: {fencer_data['Scored']}, Conceded: {fencer_data['Conceded']}, Total: {fencer_data['Total']}")
                
                # Convert matches data into a DataFrame
                match_outcomes_df = pd.DataFrame(fencer_data['Matches'])
                
                # Aggregate outcomes by opponent team
                if not match_outcomes_df.empty:
                    outcome_summary = (
                        match_outcomes_df
                        .groupby(['Opponent Team', 'Outcome'])
                        .size()
                        .unstack(fill_value=0)
                        .reindex(columns=["Victory", "Defeat", "Draw"], fill_value=0)  # Ensure all columns are present
                        .reset_index()
                    )
                    st.table(outcome_summary)
                else:
                    st.write("No match outcome data available for this fencer.")
        else:
            st.write("No match outcome data available for Qatari fencers.")

    with tabs[4]:  # Creating the new tab
        st.header("Final Rankings")

        # Extract and display final rankings
        final_rankings_df = extract_final_rankings(root, team_dict)
        if not final_rankings_df.empty:
            # Highlight Qatar in the final rankings table
            def highlight_qatar(row):
                return ['background-color: yellow' if row['Nation'] == "QAT" else '' for _ in row]

            # Display the table with Qatar highlighted
            st.dataframe(final_rankings_df.style.apply(highlight_qatar, axis=1))
        else:
            st.write("No final rankings data available.")