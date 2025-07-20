import streamlit as st
import pandas as pd
import simplekml
from io import BytesIO
import zipfile
import math

def calculate_boundary(lat, lon, radius_meters=250):
    """Menghasilkan koordinat boundary lingkaran 250m dari titik ODP"""
    boundary_points = []
    for angle in range(0, 360, 10):  # Setiap 10 derajat
        # Formula perhitungan titik boundary (approximation)
        dx = radius_meters * 0.0000089 * math.cos(math.radians(angle))
        dy = radius_meters * 0.0000089 * math.sin(math.radians(angle))
        boundary_points.append((lon + dx, lat + dy))
    return boundary_points

def create_kml_structure(kml, project_name):
    """Membuat struktur KML dengan boundary"""
    # Styles
    odp_style = simplekml.Style()
    odp_style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/ltblu-stars.png"
    odp_style.iconstyle.scale = 1.2
    
    house_style = simplekml.Style()
    house_style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/shapes/homegardenbusiness.png"
    
    boundary_style = simplekml.Style()
    boundary_style.linestyle.color = simplekml.Color.red
    boundary_style.linestyle.width = 2
    boundary_style.polystyle.color = simplekml.Color.changealphaint(50, simplekml.Color.green)

    # Folder utama
    main_folder = kml.newfolder(name=f"{project_name}.kml")
    main_folder.open = 1

    # Folder EXISTING
    existing_folder = main_folder.newfolder(name="EXISTING")
    existing_folder.open = 1
    odp_folder = existing_folder.newfolder(name="ODP")
    odp_folder.open = 1

    # Folder BOUNDARY (untuk lingkaran 250m)
    boundary_folder = main_folder.newfolder(name="BOUNDARY")
    boundary_folder.open = 1

    # Folder HOUSEHOLD
    household_folder = main_folder.newfolder(name="HOUSHOLD")
    household_folder.open = 1

    return odp_folder, household_folder, boundary_folder, odp_style, house_style, boundary_style

def main():
    st.title("Konversi Excel ke KML + Boundary 250m")
    st.write("Upload file Excel untuk generate KML dengan boundary coverage")

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
                    odp_folder, household_folder, boundary_folder, odp_style, house_style, boundary_style = create_kml_structure(kml, project_name)
                    
                    project_data = df[df['NAMA PROJECT'] == project_name]
                    
                    for _, row in project_data.iterrows():
                        # Placemark ODP
                        odp = odp_folder.newpoint(
                            name=row['ODP'],
                            description=f"Deskripsi: {row['Deskripsi']}\nProject: {row['NAMA PROJECT']}"
                        )
                        odp.coords = [(row['LONG ODP'], row['LAT ODP'])]
                        odp.style = odp_style
                        
                        # Placemark Household
                        house = household_folder.newpoint(name=row['name'])
                        house.coords = [(row['LONG PELANGGAN'], row['LAT PELANGGAN'])]
                        house.style = house_style
                        
                        # Boundary 250m dari ODP
                        boundary_points = calculate_boundary(row['LAT ODP'], row['LONG ODP'])
                        polygon = boundary_folder.newpolygon(
                            name=f"Coverage {row['ODP']}",
                            description=f"Radius 250m dari {row['ODP']} ke {row['name']}",
                            outerboundaryis=boundary_points
                        )
                        polygon.style = boundary_style
                    
                    zip_file.writestr(f"{project_name}.kml", kml.kml())
                    progress_bar.progress((i + 1) / total_projects)
            
            st.success(f"Berhasil generate {total_projects} KML dengan boundary!")
            
            # Preview contoh boundary
            with st.expander("**Contoh Visualisasi Boundary**"):
                st.image("https://i.imgur.com/J5qKZzL.png", width=400)
                st.caption("Ilustrasi boundary 250m dari ODP ke pelanggan")
            
            st.download_button(
                label="Download KML (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="projects_with_boundary.zip",
                mime="application/zip"
            )
            
        except Exception as e:
            st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
