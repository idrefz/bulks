import streamlit as st
import pandas as pd
import simplekml
from io import BytesIO
import zipfile

def main():
    st.title("Konversi Excel ke KML (Tanpa Folder)")
    st.write("Upload file Excel untuk menghasilkan KML tanpa struktur folder")
    
    # Tambahkan checkbox untuk pilihan Household
    include_household = st.checkbox("Tambah Data Household", value=True)

    uploaded_file = st.file_uploader("Pilih file Excel", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            required_columns = ['NAMA PROJECT', 'Deskripsi', 'ODP', 'LAT ODP', 'LONG ODP']
            
            # Jika Household dipilih, tambahkan validasi kolom Household
            if include_household:
                household_columns = ['name', 'LAT PELANGGAN', 'LONG PELANGGAN']
                required_columns.extend(household_columns)
            
            if not all(col in df.columns for col in required_columns):
                st.error(f"Kolom wajib: {', '.join(required_columns)}")
                if include_household:
                    st.warning("Untuk menambah Household, pastikan kolom 'name', 'LAT PELANGGAN', 'LONG PELANGGAN' tersedia")
                return

            zip_buffer = BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                progress_bar = st.progress(0)
                projects = df['NAMA PROJECT'].unique()
                total_projects = len(projects)
                
                for i, project_name in enumerate(projects):
                    kml = simplekml.Kml()
                    
                    # Style untuk ODP
                    odp_style = simplekml.Style()
                    odp_style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/ltblu-stars.png"
                    odp_style.iconstyle.scale = 1.2
                    
                    # Style untuk Household (jika digunakan)
                    if include_household:
                        house_style = simplekml.Style()
                        house_style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/shapes/homegardenbusiness.png"
                    
                    # Isi data
                    project_data = df[df['NAMA PROJECT'] == project_name]
                    
                    odp_count = 0
                    household_count = 0
                    
                    for _, row in project_data.iterrows():
                        # Point ODP
                        odp_point = kml.newpoint(
                            name=row['ODP'],
                            description=f"Deskripsi: {row['Deskripsi']}\n\nProject: {row['NAMA PROJECT']}\n\nType: ODP"
                        )
                        odp_point.coords = [(row['LONG ODP'], row['LAT ODP'])]
                        odp_point.style = odp_style
                        odp_count += 1
                        
                        # Point Household (hanya jika dipilih)
                        if include_household:
                            house_point = kml.newpoint(
                                name=row['name'],
                                description=f"Project: {row['NAMA PROJECT']}\n\nType: Household"
                            )
                            house_point.coords = [(row['LONG PELANGGAN'], row['LAT PELANGGAN'])]
                            house_point.style = house_style
                            household_count += 1
                    
                    zip_file.writestr(f"{project_name}.kml", kml.kml())
                    progress_bar.progress((i + 1) / total_projects)
                    
                    # Tampilkan info per project
                    st.info(f"Project {project_name}: {odp_count} ODP, {household_count} Household")
            
            st.success(f"Berhasil membuat {total_projects} file KML!")
            
            # Info ringkasan
            if include_household:
                st.info("KML mencakup data ODP dan Household")
            else:
                st.info("KML hanya mencakup data ODP (tanpa Household)")
            
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
