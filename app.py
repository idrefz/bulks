import streamlit as st
import pandas as pd
import simplekml
from io import BytesIO
import zipfile

def create_kml_structure(kml):
    """Membuat struktur folder lengkap meski kosong"""
    # Folder EXISTING dan subfoldernya
    existing = kml.newfolder(name="EXISTING")
    existing.newfolder(name="ODP")
    existing.newfolder(name="TIANG")
    existing.newfolder(name="DISTRIBUSI")
    existing.newfolder(name="BOUNDARY")
    existing.newfolder(name="ODC")
    existing.newfolder(name="CLOSURE")
    existing.newfolder(name="FEEDER")
    
    # Folder NEW PLANNING dan subfoldernya
    new_planning = kml.newfolder(name="NEW PLANNING")
    new_planning.newfolder(name="ODP")
    new_planning.newfolder(name="TIANG")
    new_planning.newfolder(name="DISTRIBUSI")
    new_planning.newfolder(name="BOUNDARY")
    new_planning.newfolder(name="ODC")
    new_planning.newfolder(name="FEEDER")
    new_planning.newfolder(name="CLOSURE")
    
    # Folder HOUSEHOLD
    kml.newfolder(name="HOUSEHOLD")

def main():
    st.title("Konversi Excel ke KML per Project")
    st.write("Upload file Excel untuk menghasilkan KML terpisah per project")

    uploaded_file = st.file_uploader("Pilih file Excel", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            required_columns = ['NAMA PROJECT', 'ODP', 'LAT ODP', 'LONG ODP', 'name', 'LAT PELANGGAN', 'LONG PELANGGAN']
            
            if not all(col in df.columns for col in required_columns):
                st.error(f"Kolom wajib tidak ditemukan. Pastikan file memiliki kolom: {', '.join(required_columns)}")
                return

            zip_buffer = BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                progress_bar = st.progress(0)
                projects = df['NAMA PROJECT'].unique()
                total_projects = len(projects)
                
                for i, project_name in enumerate(projects):
                    kml = simplekml.Kml()
                    create_kml_structure(kml)  # Buat struktur dasar
                    
                    # Filter data untuk project ini
                    project_data = df[df['NAMA PROJECT'] == project_name]
                    
                    # Isi data ODP ke folder EXISTING > ODP
                    odp_folder = kml.document.folder[0].folder[0]  # EXISTING > ODP
                    for _, row in project_data.iterrows():
                        odp_folder.newpoint(
                            name=row['ODP'],
                            coords=[(row['LONG ODP'], row['LAT ODP'])]
                        )
                    
                    # Isi data pelanggan ke HOUSEHOLD
                    household_folder = kml.document.folder[2]  # HOUSEHOLD
                    for _, row in project_data.iterrows():
                        household_folder.newpoint(
                            name=row['name'],
                            coords=[(row['LONG PELANGGAN'], row['LAT PELANGGAN'])]
                        )
                    
                    zip_file.writestr(f"{project_name}.kml", kml.kml())
                    progress_bar.progress((i + 1) / total_projects)
            
            st.success(f"Berhasil mengkonversi {total_projects} project!")
            st.download_button(
                label="Download KML Files (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="projects_kml.zip",
                mime="application/zip"
            )
            
            st.subheader("Struktur Folder KML")
            st.code("""
            EXISTING/
            ├── ODP/         (berisi placemark ODP)
            ├── TIANG/       (kosong)
            ├── DISTRIBUSI/ (kosong)
            ├── BOUNDARY/   (kosong)
            ├── ODC/        (kosong)
            ├── CLOSURE/     (kosong)
            └── FEEDER/      (kosong)
            
            NEW PLANNING/
            ├── ODP/        (kosong)
            ├── TIANG/      (kosong)
            ├── ...         (dst)
            
            HOUSEHOLD/      (berisi placemark pelanggan)
            """)
            
        except Exception as e:
            st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
