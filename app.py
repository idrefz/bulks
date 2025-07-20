import streamlit as st
import pandas as pd
import simplekml
from io import BytesIO
import zipfile

def main():
    st.title("Konversi Excel ke KML per Project")
    st.write("Upload file Excel untuk menghasilkan KML terpisah per project")

    # Upload file
    uploaded_file = st.file_uploader("Pilih file Excel", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            # Baca data Excel
            df = pd.read_excel(uploaded_file)
            
            # Validasi kolom
            required_columns = ['NAMA PROJECT', 'ODP', 'LAT ODP', 'LONG ODP', 'name', 'LAT PELANGGAN', 'LONG PELANGGAN']
            if not all(col in df.columns for col in required_columns):
                st.error(f"File Excel harus memiliki kolom: {', '.join(required_columns)}")
                return

            # Buat ZIP file di memory
            zip_buffer = BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                progress_bar = st.progress(0)
                total_projects = len(df['NAMA PROJECT'].unique())
                
                for i, (project_name, group) in enumerate(df.groupby('NAMA PROJECT')):
                    kml = simplekml.Kml()
                    
                    # Folder EXISTING > ODP
                    existing_folder = kml.newfolder(name="EXISTING")
                    odp_folder = existing_folder.newfolder(name="ODP")
                    
                    # Folder HOUSEHOLD
                    household_folder = kml.newfolder(name="HOUSEHOLD")
                    
                    for _, row in group.iterrows():
                        # Tambahkan ODP
                        odp_folder.newpoint(
                            name=row['ODP'],
                            coords=[(row['LONG ODP'], row['LAT ODP'])]
                        )
                        
                        # Tambahkan Pelanggan
                        household_folder.newpoint(
                            name=row['name'],
                            coords=[(row['LONG PELANGGAN'], row['LAT PELANGGAN'])]
                        )
                    
                    # Simpan KML ke ZIP
                    kml_data = kml.kml().encode('utf-8')
                    zip_file.writestr(f"{project_name}.kml", kml_data)
                    
                    # Update progress bar
                    progress_bar.progress((i + 1) / total_projects)
            
            # Download button
            st.success("Konversi selesai!")
            st.download_button(
                label="Download KML Files (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="projects_kml.zip",
                mime="application/zip"
            )
            
            # Tampilkan preview data
            st.subheader("Preview Data")
            st.dataframe(df.head())
            
        except Exception as e:
            st.error(f"Terjadi error: {str(e)}")

if __name__ == "__main__":
    main()
