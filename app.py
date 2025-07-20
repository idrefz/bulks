import streamlit as st
import pandas as pd
import simplekml
from io import BytesIO
import zipfile

def create_kml_structure(kml, project_name):
    """Membuat struktur KML dengan line string"""
    # Style
    odp_style = simplekml.Style()
    odp_style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/ltblu-stars.png"
    odp_style.iconstyle.scale = 1.2
    
    house_style = simplekml.Style()
    house_style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/shapes/homegardenbusiness.png"
    
    line_style = simplekml.Style()
    line_style.linestyle.color = simplekml.Color.red  # Warna garis merah
    line_style.linestyle.width = 3  # Ketebalan garis

    # Folder utama
    main_folder = kml.newfolder(name=f"{project_name}.kml")
    main_folder.open = 1

    # Folder EXISTING
    existing_folder = main_folder.newfolder(name="EXISTING")
    existing_folder.open = 1
    odp_folder = existing_folder.newfolder(name="ODP")
    odp_folder.open = 1

    # Folder NEW PLANING > DISTRIBUSI (untuk line string)
    new_planing_folder = main_folder.newfolder(name="NEW PLANING")
    distribusi_folder = new_planing_folder.newfolder(name="DISTRIBUSI")
    distribusi_folder.open = 1  # Folder dibuka secara default

    # Folder lainnya
    existing_folder.newfolder(name="TIANG").visibility = 0
    existing_folder.newfolder(name="BOUNDARY").visibility = 0
    existing_folder.newfolder(name="ODC").visibility = 0
    existing_folder.newfolder(name="CLOSURE").visibility = 0
    existing_folder.newfolder(name="FEEDER").visibility = 0

    # Folder HOUSEHOLD
    household_folder = main_folder.newfolder(name="HOUSHOLD")
    household_folder.open = 1

    return odp_folder, household_folder, distribusi_folder, odp_style, house_style, line_style

def main():
    st.title("Konversi Excel ke KML + Line String")
    st.write("Garis penghubung akan dibuat dari ODP ke Pelanggan")

    uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx", "xls"])

    if uploaded_file:
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
                    odp_folder, household_folder, distribusi_folder, odp_style, house_style, line_style = create_kml_structure(kml, project_name)
                    
                    project_data = df[df['NAMA PROJECT'] == project_name]
                    
                    for _, row in project_data.iterrows():
                        # Placemark ODP
                        odp = odp_folder.newpoint(
                            name=row['ODP'],
                            description=f"Deskripsi: {row['Deskripsi']}\nProject: {project_name}"
                        )
                        odp.coords = [(row['LONG ODP'], row['LAT ODP'])]
                        odp.style = odp_style
                        
                        # Placemark Pelanggan
                        house = household_folder.newpoint(name=row['name'])
                        house.coords = [(row['LONG PELANGGAN'], row['LAT PELANGGAN'])]
                        house.style = house_style
                        
                        # Line String ODP ke Pelanggan
                        line = distribusi_folder.newlinestring(
                            name=f"Jalur: {row['ODP']} ke {row['name']}",
                            description=f"Koneksi dari ODP ke pelanggan\nODP: {row['ODP']}\nPelanggan: {row['name']}",
                            coords=[
                                (row['LONG ODP'], row['LAT ODP']),
                                (row['LONG PELANGGAN'], row['LAT PELANGGAN'])
                            ]
                        )
                        line.style = line_style
                    
                    zip_file.writestr(f"{project_name}.kml", kml.kml())
                    progress_bar.progress((i + 1) / total_projects)
            
            st.success("✅ Konversi selesai! Garis penghubung ditambahkan di NEW PLANING > DISTRIBUSI")
            
            # Preview contoh output
            with st.expander("**Contoh Struktur KML**"):
                st.image("https://i.imgur.com/JQ6y5tN.png", caption="Contoh garis penghubung di Google Earth")
            
            st.download_button(
                label="Download KML (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="projects_with_lines.zip",
                mime="application/zip"
            )
            
        except Exception as e:
            st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
