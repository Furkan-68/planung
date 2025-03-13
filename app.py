import streamlit as st
import pandas as pd
import random
import json
import os
import io
import base64

# Page configuration
st.set_page_config(
    page_title="Mitarbeitereinsatz", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Increase sidebar width
st.markdown("""
<style>
    [data-testid="stSidebar"] {
        min-width: 450px;
        max-width: 450px;
    }
</style>
""", unsafe_allow_html=True)

# Define path for storing project data
DATA_FILE = "project_data.json"

# Function to generate random projects
def generate_random_projects(num_projects=5):
    project_types = ["Hardware", "Software", "Netzwerk", "Cloud", "KI", "Datenbank", "Security", "Mobile", "Web", "IoT"]
    projects = []
    
    for i in range(num_projects):
        # Random project name
        project_type = random.choice(project_types)
        project_name = f"{project_type}-Projekt {random.randint(1000, 9999)}"
        
        # Random quantity
        quantity = random.randint(1, 50)
        
        # Create project
        project = {
            "name": project_name,
            "quantity": quantity,
            "stations": {
                "Station 1": random.choice([True, False]),
                "Station 2": random.choice([True, False]),
                "Station 3": random.choice([True, False]),
                "Station 4": random.choice([True, False]),
                "Station 5": random.choice([True, False]),
                "Station 6": random.choice([True, False]),
                "Station 7": random.choice([True, False])
            }
        }
        projects.append(project)
    
    return projects

# Function to save projects to file
def save_projects():
    try:
        with open(DATA_FILE, "w") as f:
            json.dump(st.session_state.projects, f)
    except Exception as e:
        st.error(f"Fehler beim Speichern der Projekte: {e}")

# Function to load projects from file
def load_projects():
    try:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r") as f:
                return json.load(f)
        else:
            return generate_random_projects()
    except Exception as e:
        st.error(f"Fehler beim Laden der Projekte: {e}")
        return generate_random_projects()

# Function to export projects to Excel
def export_to_excel():
    # Convert projects to DataFrames
    projects_df = pd.DataFrame([
        {"name": p["name"], "quantity": p["quantity"]} 
        for p in st.session_state.projects
    ])
    
    # Convert stations to DataFrame
    stations_data = []
    for i, project in enumerate(st.session_state.projects):
        if 'stations' in project:
            for station, is_active in project['stations'].items():
                stations_data.append({
                    "project_index": i,
                    "project_name": project["name"],
                    "station": station,
                    "active": is_active
                })
    stations_df = pd.DataFrame(stations_data)
    
    # Create Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        projects_df.to_excel(writer, sheet_name='Projects', index=False)
        stations_df.to_excel(writer, sheet_name='Stations', index=False)
    
    # Return the Excel file as bytes
    output.seek(0)
    return output.getvalue()

# Function to create a download link for Excel
def get_excel_download_link(excel_file, file_name="projekte.xlsx"):
    b64 = base64.b64encode(excel_file).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{file_name}">Download Excel Datei</a>'
    return href

# Function to import projects from Excel
def import_from_excel(file):
    try:
        # Check if the file is readable
        if file is None:
            raise ValueError("Keine Datei ausgewählt")
            
        # Read Excel file - try both sheet_name options to be more robust
        try:
            projects_df = pd.read_excel(file, sheet_name='Projects')
        except:
            projects_df = pd.read_excel(file, sheet_name=0)
            
        # Check for required columns
        if "name" not in projects_df.columns or "quantity" not in projects_df.columns:
            # Try to use first two columns if they exist
            if len(projects_df.columns) >= 2:
                projects_df.columns = ["name", "quantity"] + list(projects_df.columns[2:])
            else:
                raise ValueError("Erforderliche Spalten 'name' und 'quantity' nicht gefunden")
        
        # Try to read stations sheet if it exists
        try:
            stations_df = pd.read_excel(file, sheet_name='Stations')
        except:
            # If Stations sheet doesn't exist, create an empty DataFrame
            stations_df = pd.DataFrame(columns=["project_index", "project_name", "station", "active"])
        
        # Convert to project format
        projects = []
        for i, row in projects_df.iterrows():
            try:
                # Ensure quantity is a valid integer
                try:
                    quantity = int(row["quantity"])
                except:
                    quantity = 1  # Default to 1 if conversion fails
                
                project = {
                    "name": str(row["name"]),
                    "quantity": quantity,
                    "stations": {
                        "Wareneingang": False,
                        "Labeln": False,
                        "Sichtkontrolle": False,
                        "Erase-Vorgang": False,
                        "Eingabe in das Kundenportal": False,
                        "Einlagerung": False,
                        "Ausgang": False
                    }
                }
                projects.append(project)
            except Exception as e:
                st.warning(f"Zeile {i+1} konnte nicht importiert werden: {e}")
        
        # Check if we have any valid projects
        if not projects:
            raise ValueError("Keine gültigen Projekte in der Datei gefunden")
        
        # Apply station settings if stations sheet exists and has required columns
        required_columns = ["project_index", "station", "active"]
        if all(col in stations_df.columns for col in required_columns):
            for _, row in stations_df.iterrows():
                try:
                    project_index = int(row["project_index"])
                    station = str(row["station"])
                    is_active = bool(row["active"])
                    
                    if project_index < len(projects) and station in projects[project_index]["stations"]:
                        projects[project_index]["stations"][station] = is_active
                except Exception as e:
                    pass  # Skip invalid station entries
        
        return projects
    except Exception as e:
        raise Exception(f"Fehler beim Importieren der Excel-Datei: {e}")
        return None

# Initialize project list in session state if not exists
if 'projects' not in st.session_state:
    # Make sure we generate new random projects if no data file exists
    initial_projects = load_projects()
    if not initial_projects or len(initial_projects) == 0:
        initial_projects = generate_random_projects()
    st.session_state.projects = initial_projects
    
# Save projects when app state changes
def on_change():
    save_projects()
    
st.session_state.on_change = on_change

# Initialize selected project index if not exists
if 'selected_project_index' not in st.session_state:
    st.session_state.selected_project_index = 0 if st.session_state.projects else None

# Sidebar for project management
with st.sidebar:
    # Project list display first
    if st.session_state.projects:
        st.subheader("Liste der Projekte")
        
        for i, project in enumerate(st.session_state.projects):
            col1, col2, col3, col4 = st.columns([5, 2, 1, 1])
            
            with col1:
                # Make project name clickable for selection
                if st.button(
                    project['name'], 
                    key=f"select_{i}",
                    use_container_width=True,
                    type="secondary"
                ):
                    st.session_state.selected_project_index = i
                    st.rerun()
            
            with col2:
                # Update quantity for each project
                new_qty = st.number_input(
                    "Anzahl", 
                    min_value=1, 
                    value=project["quantity"], 
                    step=1,
                    key=f"qty_{i}"
                )
                
                if new_qty != project["quantity"]:
                    project["quantity"] = new_qty
                    save_projects()
            
            with col3:
                # Settings button for each project
                if st.button("⚙️", key=f"settings_{i}", help="Projekteinstellungen"):
                    st.session_state.selected_project_index = i
                    st.session_state.show_settings_dialog = True
                    st.rerun()
            
            with col4:
                # Delete button for each project
                if st.button("🗑️", key=f"delete_{i}", help="Projekt löschen"):
                    if i == st.session_state.selected_project_index:
                        # If deleting selected project, select first one or None
                        if len(st.session_state.projects) > 1:
                            st.session_state.selected_project_index = 0
                        else:
                            st.session_state.selected_project_index = None
                    elif i < st.session_state.selected_project_index:
                        # If deleting a project before the selected one, adjust the index
                        st.session_state.selected_project_index -= 1
                    
                    st.session_state.projects.pop(i)
                    save_projects()
                    st.rerun()
        
        # Add Calculate button under the project list
        st.divider()
        any_stations_selected = any(
            sum(1 for value in p.get('stations', {}).values() if value) > 0 
            for p in st.session_state.projects
        )
        
        if any_stations_selected:
            if st.button("🧮 Berechnen", use_container_width=True, key="calculate_button"):
                st.session_state.show_results = True
                st.rerun()
        else:
            st.warning("Bitte wählen Sie mindestens eine Station über den ⚙️ Einstellungen-Button bei einem Projekt aus.")
    else:
        st.warning("Keine Projekte vorhanden.")
        st.session_state.selected_project_index = None
    
    # Project management interface (now after project list)
    st.divider()
    st.subheader("Projekt hinzufügen")
    
    # Project name input
    new_projekt_name = st.text_input("Projektname:")
    
    # Quantity input
    new_projekt_anzahl = st.number_input("Anzahl:", min_value=1, value=1, step=1)
    
    # Add Project and Settings buttons
    col1, col2 = st.columns([1, 1])
    
    with col1:
        # Add new project with "+" button
        if st.button("➕ Hinzufügen", help="Neues Projekt hinzufügen", use_container_width=True) and new_projekt_name:
            # Initialize stations for new project
            new_project = {
                "name": new_projekt_name, 
                "quantity": new_projekt_anzahl,
                "stations": {
                    "Station 1": False,
                    "Station 2": False,
                    "Station 3": False,
                    "Station 4": False,
                    "Station 5": False,
                    "Station 6": False,
                    "Station 7": False
                }
            }
            st.session_state.projects.append(new_project)
            # Set the newly added project as selected
            st.session_state.selected_project_index = len(st.session_state.projects) - 1
            save_projects()
            st.rerun()
    
    with col2:
        # Settings button for new project
        if st.button("⚙️ Einstellungen", help="Stationen für neues Projekt konfigurieren", use_container_width=True):
            # Create a temporary project for configuration
            if 'temp_project' not in st.session_state:
                st.session_state.temp_project = {
                    "name": new_projekt_name if new_projekt_name else "Neues Projekt",
                    "quantity": new_projekt_anzahl,
                    "stations": {
                        "Station 1": False,
                        "Station 2": False,
                        "Station 3": False,
                        "Station 4": False,
                        "Station 5": False,
                        "Station 6": False,
                        "Station 7": False
                    }
                }
            st.session_state.temp_project_settings = True
            st.rerun()
    
    # Project settings are now shown in a dialog when the settings button is clicked

if 'show_results' not in st.session_state:
    st.session_state.show_results = False

# Initialize dialog state if needed
if 'show_settings_dialog' not in st.session_state:
    st.session_state.show_settings_dialog = False

# Main content area
st.title("Mitarbeitereinsatz")

# Define dialog function for project settings
@st.dialog("Projekteinstellungen")
def show_project_settings(project_index):
    if project_index is not None and project_index < len(st.session_state.projects):
        selected_project = st.session_state.projects[project_index]
        
        # Initialize project-specific toggle values if not exists
        if 'stations' not in selected_project:
            selected_project['stations'] = {
                "Station 1": False,
                "Station 2": False,
                "Station 3": False,
                "Station 4": False,
                "Station 5": False,
                "Station 6": False,
                "Station 7": False
            }
        
        # Dialog content
        st.subheader(f"Stationen für {selected_project['name']}")
        
        # Create two columns for station checkboxes
        col1, col2 = st.columns(2)
        
        # Split stations into two groups
        stations = list(selected_project['stations'].keys())
        half = len(stations) // 2 + len(stations) % 2
        
        # First column
        with col1:
            for station in stations[:half]:
                value = st.checkbox(
                    station, 
                    value=selected_project['stations'][station],
                    key=f"dialog_{station}_{project_index}"
                )
                if value != selected_project['stations'][station]:
                    selected_project['stations'][station] = value
                    save_projects()
        
        # Second column
        with col2:
            for station in stations[half:]:
                value = st.checkbox(
                    station, 
                    value=selected_project['stations'][station],
                    key=f"dialog_{station}_{project_index}"
                )
                if value != selected_project['stations'][station]:
                    selected_project['stations'][station] = value
                    save_projects()
        
        # Close button
        if st.button("Schließen", key="close_settings"):
            st.session_state.show_settings_dialog = False
            st.rerun()

# Dialog for temporary project settings
@st.dialog("Einstellungen für neues Projekt")
def show_temp_project_settings():
    if 'temp_project' in st.session_state:
        # Dialog content
        st.subheader(f"Stationen für neues Projekt")
        
        # Create two columns for station checkboxes
        col1, col2 = st.columns(2)
        
        # Split stations into two groups
        stations = list(st.session_state.temp_project['stations'].keys())
        half = len(stations) // 2 + len(stations) % 2
        
        # First column
        with col1:
            for station in stations[:half]:
                value = st.checkbox(
                    station, 
                    value=st.session_state.temp_project['stations'][station],
                    key=f"temp_dialog_{station}"
                )
                if value != st.session_state.temp_project['stations'][station]:
                    st.session_state.temp_project['stations'][station] = value
        
        # Second column
        with col2:
            for station in stations[half:]:
                value = st.checkbox(
                    station, 
                    value=st.session_state.temp_project['stations'][station],
                    key=f"temp_dialog_{station}"
                )
                if value != st.session_state.temp_project['stations'][station]:
                    st.session_state.temp_project['stations'][station] = value
        
        # Buttons
        col1, col2 = st.columns(2)
        
        with col1:
            # Apply and save button
            if st.button("Speichern", key="save_temp_project"):
                # Create new project with selected stations
                if 'name' in st.session_state.temp_project and st.session_state.temp_project['name']:
                    st.session_state.projects.append(st.session_state.temp_project)
                    st.session_state.selected_project_index = len(st.session_state.projects) - 1
                    save_projects()
                    st.session_state.temp_project_settings = False
                    del st.session_state.temp_project
                    st.rerun()
                else:
                    st.error("Bitte geben Sie einen Projektnamen ein.")
        
        with col2:
            # Close button
            if st.button("Abbrechen", key="close_temp_settings"):
                st.session_state.temp_project_settings = False
                del st.session_state.temp_project
                st.rerun()

# Show dialogs if needed
if st.session_state.selected_project_index is not None and st.session_state.show_settings_dialog:
    show_project_settings(st.session_state.selected_project_index)

if 'temp_project_settings' in st.session_state and st.session_state.temp_project_settings:
    show_temp_project_settings()

# Add Excel Import/Export at the end of the sidebar
with st.sidebar:
    # Add a divider to separate from project management
    st.divider()
    
    # Excel import/export section
    st.subheader("Excel Import/Export")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        # Excel export button
        if st.button("📤 Export", help="Als Excel exportieren"):
            excel_data = export_to_excel()
            st.markdown(get_excel_download_link(excel_data), unsafe_allow_html=True)
    
    with col2:
        # Excel import button
        uploaded_file = st.file_uploader("📥 Import", type=["xlsx"], help="Excel-Datei importieren", key="excel_uploader")
        
        # Add import button to control when import happens (instead of automatic on upload)
        if uploaded_file is not None and st.button("Importieren", key="import_button"):
            try:
                # Add debug information
                st.info("Importiere Excel-Datei...")
                
                # Import projects
                imported_projects = import_from_excel(uploaded_file)
                
                if imported_projects and len(imported_projects) > 0:
                    st.session_state.projects = imported_projects
                    st.session_state.selected_project_index = 0
                    save_projects()
                    st.success("Projekte importiert!", icon="✅")
                    st.rerun()
                else:
                    st.error("Keine gültigen Projekte in der Excel-Datei gefunden.")
            except Exception as e:
                st.error(f"Fehler beim Import: {str(e)}")

if st.session_state.projects:
    # Display welcome message when no results are shown yet
    if not st.session_state.show_results:
        st.info("Klicken Sie auf den 'Rechnen' Button in der Seitenleiste, um die Berechnungsergebnisse anzuzeigen.")
else:
    st.info("Keine Projekte vorhanden. Bitte fügen Sie im Seitenmenü ein Projekt hinzu.")

# Results display
if st.session_state.show_results:
    # Display current projects first
    st.subheader("Aktuelle Projekte")
    
    # Create a dataframe to display projects
    projects_data = []
    for i, p in enumerate(st.session_state.projects):
        # Get active stations for this project
        active_station_names = []
        if 'stations' in p:
            active_station_names = [station for station, is_active in p['stations'].items() if is_active]
            active_stations_count = len(active_station_names)
        else:
            active_stations_count = 0
        
        # Format active stations as comma-separated list
        active_stations_text = ", ".join(active_station_names) if active_station_names else "Keine"
            
        projects_data.append({
            "Projekt": p["name"],
            "Anzahl": p["quantity"],
            "Aktive Stationen": active_stations_count,
            "Ausgewählte Stationen": active_stations_text
        })
    
    projects_df = pd.DataFrame(projects_data)
    st.dataframe(projects_df)
    
    # Total quantity
    total_quantity = sum(p["quantity"] for p in st.session_state.projects)
    st.write(f"Gesamtanzahl: **{total_quantity}**")
    
    # Calculation results
    st.subheader("Berechnungsergebnisse")
    
    # Import for random number generation
    import random
    
    # Get all unique active stations across all projects
    all_stations = set()
    for project in st.session_state.projects:
        if 'stations' in project:
            for station, is_active in project['stations'].items():
                if is_active:
                    all_stations.add(station)
    
    if not all_stations:
        st.warning("Keine Stationen ausgewählt. Wählen Sie im Seitenmenü für mindestens ein Projekt Stationen aus.")
    else:
        # Create results for each station
        station_results = []
        
        for station in all_stations:
            # Generate random number of employees (1-3)
            mitarbeiter = random.randint(1, 3)
            
            station_results.append({
                "Station": station,
                "Anzahl Mitarbeiter": mitarbeiter
            })
        
        # Create DataFrame and display table
        results_df = pd.DataFrame(station_results)
        st.table(results_df)
        
        # Summary statistics
        st.subheader("Zusammenfassung")
        total_mitarbeiter = results_df["Anzahl Mitarbeiter"].sum()
        total_stations = len(station_results)
        
        st.write(f"Anzahl Stationen: **{total_stations}**")
        st.write(f"Benötigte Mitarbeiter insgesamt: **{total_mitarbeiter}**")
        
        # No reset button needed anymore