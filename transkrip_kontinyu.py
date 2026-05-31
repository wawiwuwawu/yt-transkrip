import yt_dlp
import whisper
import pandas as pd
import os
import torch
import re
from datetime import timedelta

def download_audio_ytdlp(video_url, output_path="temp_audio"):
    """Download audio menggunakan yt-dlp"""
    ydl_opts = {
        'format': 'bestaudio/best',
        'extractaudio': True,
        'audioformat': 'mp3',
        'outtmpl': f'{output_path}.%(ext)s',
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }],
        'quiet': True,
    }
    
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            # Get video info
            info = ydl.extract_info(video_url, download=False)
            video_info = {
                'title': info.get('title', 'Unknown'),
                'duration': info.get('duration', 0),
                'uploader': info.get('uploader', 'Unknown'),
            }
            
            # Download audio
            ydl.download([video_url])
            
            # Find the downloaded file
            audio_file = None
            for ext in ['mp3', 'wav', 'm4a', 'webm']:
                potential_file = f"{output_path}.{ext}"
                if os.path.exists(potential_file):
                    audio_file = potential_file
                    break
            
            if not audio_file:
                raise Exception("Audio file tidak ditemukan setelah download")
                
            return audio_file, video_info
            
    except Exception as e:
        print(f"Error saat download: {e}")
        return None, None

def buat_transkrip_kontinyu(video_url, output_name="transkrip_kontinyu"):
    """
    Membuat transkrip kontinyu tanpa timestamp
    """
    try:
        print("🎬 YOUTUBE TRANSCRIPT GENERATOR - KONTINYU MODE")
        print("=" * 55)
        
        # --- STEP 1: DOWNLOAD AUDIO ---
        print(f"\n📥 Mengunduh audio dari YouTube...")
        print(f"🔗 URL: {video_url}")
        
        audio_file, video_info = download_audio_ytdlp(video_url, "temp_audio")
        
        if not audio_file or not video_info:
            raise Exception("Gagal mengunduh audio dari YouTube")
        
        print(f"✅ Audio berhasil diunduh!")
        print(f"📹 Judul: {video_info['title']}")
        print(f"👤 Channel: {video_info['uploader']}")
        
        # --- STEP 2: LOAD AI MODEL ---
        print(f"\n🤖 Memuat model Whisper AI...")
        model = whisper.load_model("base")
        device = "cuda" if torch.cuda.is_available() else "cpu"
        print(f"🖥️ Device: {device}")
        
        # --- STEP 3: TRANSCRIBE ---
        print(f"\n🎙️ Memproses transkrip...")
        print("⏳ Sedang menganalisis audio...")
        
        result = model.transcribe(audio_file, fp16=(device=="cuda"))
        segments = result['segments']
        
        print(f"✅ Transkrip selesai!")
        print(f"📊 Total segmen: {len(segments)}")
        
        # Clean up
        if os.path.exists(audio_file):
            os.remove(audio_file)
        
        # --- STEP 4: CREATE CONTINUOUS TEXT ---
        print(f"\n📝 Membuat transkrip kontinyu...")
        
        # Gabungkan semua teks menjadi satu
        full_text = ""
        for seg in segments:
            text = seg['text'].strip()
            # Tambahkan spasi jika diperlukan
            if full_text and not full_text.endswith(('.', '!', '?', ',', ' ')):
                full_text += " "
            full_text += text
        
        # Bersihkan teks
        full_text = ' '.join(full_text.split())
        
        # Perbaiki tanda baca dan spasi
        full_text = re.sub(r'\s+([,.!?])', r'\\1', full_text)  # Hapus spasi sebelum tanda baca
        full_text = re.sub(r'([.!?])\s*', r'\\1 ', full_text)   # Pastikan ada spasi setelah tanda baca
        
        # --- SAVE FILES ---
        
        # 1. Text file kontinyu
        txt_filename = f"{output_name}.txt"
        with open(txt_filename, 'w', encoding='utf-8') as f:
            f.write("TRANSKRIP KONTINYU\n")
            f.write("=" * 50 + "\n")
            f.write(f"Video: {video_info['title']}\n")
            f.write(f"Channel: {video_info['uploader']}\n")
            f.write(f"Tanggal: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n\n")
            
            # Bagi menjadi paragraf untuk readability
            sentences = full_text.replace('. ', '.\\n').split('\\n')
            current_paragraph = ""
            
            for sentence in sentences:
                sentence = sentence.strip()
                if not sentence:
                    continue
                
                # Jika paragraf akan terlalu panjang, buat paragraf baru
                if current_paragraph and len(current_paragraph + " " + sentence) > 300:
                    f.write(current_paragraph.strip() + "\\n\\n")
                    current_paragraph = sentence
                else:
                    if current_paragraph:
                        current_paragraph += " " + sentence
                    else:
                        current_paragraph = sentence
            
            # Tulis paragraf terakhir
            if current_paragraph:
                f.write(current_paragraph.strip())
        
        # 2. Excel dengan transkrip kontinyu
        excel_filename = f"{output_name}.xlsx"
        
        # Data untuk Excel
        excel_data = {
            'Informasi': [
                'Judul Video',
                'Channel',
                'Total Segmen',
                'Total Kata',
                'Total Karakter',
                'Tanggal Transkrip',
                'Transkrip Lengkap'
            ],
            'Detail': [
                video_info['title'],
                video_info['uploader'],
                len(segments),
                len(full_text.split()),
                len(full_text),
                pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S'),
                full_text
            ]
        }
        
        df = pd.DataFrame(excel_data)
        
        with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Transkrip Kontinyu', index=False)
            
            # Format worksheet
            worksheet = writer.sheets['Transkrip Kontinyu']
            worksheet.column_dimensions['A'].width = 20
            worksheet.column_dimensions['B'].width = 100
            
            # Set tinggi baris untuk transkrip lengkap (baris terakhir)
            worksheet.row_dimensions[8].height = 200
            
            # Enable text wrapping untuk kolom B
            from openpyxl.styles import Alignment
            for row in worksheet.iter_rows(min_row=2, max_row=8, min_col=2, max_col=2):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
        
        # 3. Plain text file (hanya transkrip, tanpa header)
        plain_filename = f"{output_name}_plain.txt"
        with open(plain_filename, 'w', encoding='utf-8') as f:
            f.write(full_text)
        
        # 4. Word-friendly format
        word_filename = f"{output_name}_word_format.txt"
        with open(word_filename, 'w', encoding='utf-8') as f:
            # Format untuk easy copy-paste ke Word
            sentences = full_text.replace('. ', '.\\n').split('\\n')
            paragraph_count = 1
            current_paragraph = ""
            
            for sentence in sentences:
                sentence = sentence.strip()
                if not sentence:
                    continue
                
                if current_paragraph and len(current_paragraph + " " + sentence) > 250:
                    f.write(f"Paragraf {paragraph_count}:\\n")
                    f.write(current_paragraph.strip() + "\\n\\n")
                    paragraph_count += 1
                    current_paragraph = sentence
                else:
                    if current_paragraph:
                        current_paragraph += " " + sentence
                    else:
                        current_paragraph = sentence
            
            if current_paragraph:
                f.write(f"Paragraf {paragraph_count}:\\n")
                f.write(current_paragraph.strip())
        
        # Calculate statistics
        total_words = len(full_text.split())
        total_chars = len(full_text)
        
        print("\n🎉 SELESAI! File transkrip kontinyu berhasil dibuat:")
        print(f"📝 Text Kontinyu        : {txt_filename}")
        print(f"📊 Excel Format        : {excel_filename}")
        print(f"📄 Plain Text          : {plain_filename}")
        print(f"📋 Word Format         : {word_filename}")
        print("\n📈 STATISTIK:")
        print(f"   • Total kata      : {total_words:,}")
        print(f"   • Total karakter  : {total_chars:,}")
        print(f"   • Total segmen    : {len(segments)}")
        print("=" * 55)
        
        # Preview teks (100 karakter pertama)
        preview = full_text[:100] + "..." if len(full_text) > 100 else full_text
        print(f"\\n📖 PREVIEW TRANSKRIP:")
        print(f'"{preview}"')
        
        return {
            'status': 'success',
            'full_text': full_text,
            'files': {
                'txt': txt_filename,
                'excel': excel_filename,
                'plain': plain_filename,
                'word': word_filename
            },
            'stats': {
                'words': total_words,
                'characters': total_chars,
                'segments': len(segments)
            }
        }
        
    except Exception as e:
        print(f"❌ ERROR: {e}")
        import traceback
        traceback.print_exc()
        return {'status': 'error', 'message': str(e)}

# --- MAIN EXECUTION ---
if __name__ == "__main__":
    print("🚀 Memulai transkrip kontinyu...")
    
    # URL video
    video_url = "https://youtu.be/l3X2SaCndiI?si=CfkAJHcFj0y254Pr"
    
    # Output name
    output_name = "transkrip_kontinyu_hasil"
    
    print(f"📹 URL: {video_url}")
    print(f"📄 Output: {output_name}")
    print("-" * 55)
    
    # Process
    result = buat_transkrip_kontinyu(video_url, output_name)
    
    if result['status'] == 'success':
        print(f"\\n✨ BERHASIL!")
        print(f"Transkrip kontinyu siap digunakan.")
        print(f"📁 File tersimpan di folder ini.")
    else:
        print(f"\\n❌ GAGAL: {result['message']}")
        print("💡 Pastikan URL valid dan koneksi internet stabil.")
