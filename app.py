import streamlit as st
import pandas as pd
import simplekml
from io import BytesIO
import zipfile

def main():
    st.title("Konversi Excel ke KML dengan Ikon Khusus")
    st.write("Upload file Excel untuk menghasilkan KML dengan ikon custom")

    # URL ikon dari Google Maps
    ODP_ICON = "http://maps.google.com/mapfiles/kml/paddle/ltblu-stars.png"
    HOUSE_ICON = "http://maps.google.com/mapfiles/kml/shapes/homegardenbusiness.png"

    uploaded_file = st.file_uploader("Pilih file Excel", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            required_columns = ['NAMA PROJECT', 'ODP', 'LAT ODP', 'LONG ODP', 'name', 'LAT PELANGGAN', 'LONG PELANGGAN']
            
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
                    
                    # Isi data ODP dengan ikon khusus
                    for _, row in project_data.iterrows():
                        # Placemark ODP
                        odp = odp_folder.newpoint(name=row['ODP'])
                        odp.coords = [(row['LONG ODP'], row['LAT ODP'])]
                        odp.style.iconstyle.icon.href = ODP_ICON
                        odp.style.iconstyle.scale = 1.2  # Sesuaikan ukuran ikon
                        
                        # Placemark Pelanggan
                        house = household_folder.newpoint(name=row['name'])
                        house.coords = [(row['LONG PELANGGAN'], row['LAT PELANGGAN'])]
                        house.style.iconstyle.icon.href = HOUSE_ICON
                        house.style.iconstyle.scale = 1.0
                    
                    zip_file.writestr(f"{project_name}.kml", kml.kml())
                    progress_bar.progress((i + 1) / total_projects)
            
            st.success(f"Berhasil mengkonversi {total_projects} project!")
            
            # Tampilkan preview ikon
            col1, col2 = st.columns(2)
            with col1:
                st.image(ODP_ICON, caption="Ikon ODP", width=50)
            with col2:
                st.image(HOUSE_ICON, caption="Ikon Household", width=50)
            
            st.download_button(
                label="Download KML Files (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="projects_kml.zip",
                mime="application/zip"
            )
            
        except Exception as e:
            st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
