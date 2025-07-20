import streamlit as st
import pandas as pd
import simplekml
from io import BytesIO
import zipfile

def main():
    st.title("Konversi Excel ke KML dengan Deskripsi ODP")
    st.write("Upload file Excel untuk menghasilkan KML dengan deskripsi lengkap")

    # URL ikon dari Google Maps
    ODP_ICON = "http://maps.google.com/mapfiles/kml/paddle/ltblu-stars.png"
    HOUSE_ICON = "http://maps.google.com/mapfiles/kml/shapes/homegardenbusiness.png"

    uploaded_file = st.file_uploader("Pilih file Excel", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            required_columns = ['NAMA PROJECT', 'Deskripsi', 'ODP', 'LAT ODP', 'LONG ODP', 
                              'name', 'LAT PELANGGAN', 'LONG PELANGGAN']
            
            if not all(col in df.columns for col in required_columns):
                st.error(f"File Excel harus memiliki kolom: {', '.join(required_columns)}")
                return

            zip_buffer = BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                progress_bar = st.progress(0)
                projects = df['NAMA PROJECT'].unique()
                total_projects = len(projects)
                
                for i, project_name in enumerate(projects):
                    kml = simplekml.Kml()
                    
                    # Folder EXISTING > ODP
                    existing_folder = kml.newfolder(name="EXISTING")
                    odp_folder = existing_folder.newfolder(name="ODP")
                    
                    # Folder HOUSEHOLD
                    household_folder = kml.newfolder(name="HOUSEHOLD")

                    # Filter data untuk project ini
                    project_data = df[df['NAMA PROJECT'] == project_name]
                    
                    # Isi data ODP dengan deskripsi
                    for _, row in project_data.iterrows():
                        # Placemark ODP dengan deskripsi
                        odp = odp_folder.newpoint(
                            name=row['ODP'],
                            description=f"Deskripsi: {row['Deskripsi']}\n\nProject: {row['NAMA PROJECT']}"
                        )
                        odp.coords = [(row['LONG ODP'], row['LAT ODP'])]
                        odp.style.iconstyle.icon.href = ODP_ICON
                        odp.style.iconstyle.scale = 1.2
                        
                        # Placemark Pelanggan
                        house = household_folder.newpoint(name=row['name'])
                        house.coords = [(row['LONG PELANGGAN'], row['LAT PELANGGAN'])]
                        house.style.iconstyle.icon.href = HOUSE_ICON
                        house.style.iconstyle.scale = 1.0
                    
                    zip_file.writestr(f"{project_name}.kml", kml.kml())
                    progress_bar.progress((i + 1) / total_projects)
            
            st.success(f"Berhasil mengkonversi {total_projects} project!")
            
            # Tampilkan contoh output
            with st.expander("Contoh Struktur KML"):
                st.code("""
                <Placemark>
                    <name>ODP-IBU-FAN/036</name>
                    <description>
                        Deskripsi: EXPAND 1:8
                        Project: BNT,IBU_PT2_SUDRAJAT WOO15920444,VA EXPAND
                    </description>
                    <Point>
                        <coordinates>105.855777,-6.49482</coordinates>
                    </Point>
                    <Style>
                        <IconStyle>
                            <Icon>
                                <href>http://maps.google.com/mapfiles/kml/paddle/ltblu-stars.png</href>
                            </Icon>
                        </IconStyle>
                    </Style>
                </Placemark>
                """, language='xml')
            
            st.download_button(
                label="Download KML Files (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="projects_with_descriptions.zip",
                mime="application/zip"
            )
            
        except Exception as e:
            st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
