import streamlit as st
import pandas as pd
import simplekml
from io import BytesIO
import zipfile

def create_kml_structure(kml, project_name):
    """Membuat struktur KML persis seperti contoh"""
    # Style untuk ODP
    odp_style = simplekml.Style()
    odp_style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/ltblu-stars.png"
    odp_style.iconstyle.scale = 1.2
    
    # Style untuk Household
    house_style = simplekml.Style()
    house_style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/shapes/homegardenbusiness.png"

    # Folder utama
    main_folder = kml.newfolder(name=f"{project_name}.kml")
    main_folder.open = 1

    # Folder EXISTING
    existing_folder = main_folder.newfolder(name="EXISTING")
    existing_folder.open = 1

    # Subfolder EXISTING
    odp_folder = existing_folder.newfolder(name="ODP")
    odp_folder.open = 1
    
    existing_folder.newfolder(name="TIANG").visibility = 0
    existing_folder.newfolder(name="DISTRIBUSI").visibility = 0
    existing_folder.newfolder(name="BOUNDARY").visibility = 0
    existing_folder.newfolder(name="ODC").visibility = 0
    existing_folder.newfolder(name="CLOSURE").visibility = 0
    existing_folder.newfolder(name="FEEDER").visibility = 0

    # Folder NEW PLANING (bukan PLANNING)
    new_planing_folder = main_folder.newfolder(name="NEW PLANING")
    new_planing_folder.visibility = 0
    
    new_planing_folder.newfolder(name="ODP").visibility = 0
    new_planing_folder.newfolder(name="TIANG").visibility = 0
    new_planing_folder.newfolder(name="DISTRIBUSI").visibility = 0
    
    boundary_folder = new_planing_folder.newfolder(name="BOUNDARY")
    boundary_folder.visibility = 0
    boundary_folder.open = 1
    
    new_planing_folder.newfolder(name="ODC").visibility = 0
    
    feeder_folder = new_planing_folder.newfolder(name="FEEDER")
    feeder_folder.visibility = 0
    feeder_folder.open = 1
    
    closure_folder = new_planing_folder.newfolder(name="CLOSURE")
    closure_folder.visibility = 0
    closure_folder.open = 1

    # Folder BOUNDARY tambahan
    main_folder.newfolder(name="BOUNDARY").visibility = 0

    # Folder HOUSHOLD (bukan HOUSEHOLD)
    household_folder = main_folder.newfolder(name="HOUSHOLD")
    household_folder.open = 1

    return odp_folder, household_folder, odp_style, house_style

def main():
    st.title("Konversi Excel ke KML (Struktur Presisi)")
    st.write("Upload file Excel untuk menghasilkan KML dengan struktur spesifik")

    uploaded_file = st.file_uploader("Pilih file Excel", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            required_columns = ['NAMA PROJECT', 'Deskripsi', 'ODP', 'LAT ODP', 'LONG ODP', 
                              'name', 'LAT PELANGGAN', 'LONG PELANGGAN']
            
            if not all(col in df.columns for col in required_columns):
                st.error(f"Kolom wajib: {', '.join(required_columns)}")
                return

            zip_buffer = BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                progress_bar = st.progress(0)
                projects = df['NAMA PROJECT'].unique()
                total_projects = len(projects)
                
                for i, project_name in enumerate(projects):
                    kml = simplekml.Kml()
                    odp_folder, household_folder, odp_style, house_style = create_kml_structure(kml, project_name)
                    
                    # Isi data
                    project_data = df[df['NAMA PROJECT'] == project_name]
                    
                    for _, row in project_data.iterrows():
                        # Placemark ODP
                        odp = odp_folder.newpoint(
                            name=row['ODP'],
                            description=f"Deskripsi: {row['Deskripsi']}\n\nProject: {row['NAMA PROJECT']}"
                        )
                        odp.coords = [(row['LONG ODP'], row['LAT ODP'])]
                        odp.style = odp_style
                        odp.open = 1
                        
                        # Placemark Household
                        house = household_folder.newpoint(name=row['name'])
                        house.coords = [(row['LONG PELANGGAN'], row['LAT PELANGGAN'])]
                        house.style = house_style
                        house.open = 1
                    
                    zip_file.writestr(f"{project_name}.kml", kml.kml())
                    progress_bar.progress((i + 1) / total_projects)
            
            st.success(f"Berhasil membuat {total_projects} file KML!")
            
            st.download_button(
                label="Download KML (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="projects_kml.zip",
                mime="application/zip"
            )
            
        except Exception as e:
            st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
