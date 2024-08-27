import os

def get_new_output_file(base_path, counter):
    return f"{base_path}_{counter}.txt"

def merge_text_files(input_folder, output_base_path, max_size_bytes=4 * 1024 * 1024):  # 10 MB
    file_counter = 1
    current_size = 0
    output_file = get_new_output_file(output_base_path, file_counter)
    
    with open(output_file, 'w', encoding='utf-8') as outfile:
        for filename in os.listdir(input_folder):
            if filename.endswith(".txt"):
                file_path = os.path.join(input_folder, filename)
                with open(file_path, 'r', encoding='utf-8') as infile:
                    content = infile.read()
                    content_size = len(content.encode('utf-8'))
                    
                    if current_size + content_size > max_size_bytes:
                        # Close current file and open a new one
                        outfile.close()
                        file_counter += 1
                        output_file = get_new_output_file(output_base_path, file_counter)
                        outfile = open(output_file, 'w', encoding='utf-8')
                        current_size = 0
                    
                    outfile.write(content)
                    outfile.write("\n\n--- Next Email ---\n\n")
                    current_size += content_size + len("\n\n--- Next Email ---\n\n".encode('utf-8'))
                
                print(f"Processed: {filename}", end='\r')

    print(f"\nMerge completed. Created {file_counter} file(s).")

if __name__ == "__main__":
    # Adjust this line
    input_folder = r"C:\path\to\your\input\folder"
    output_base_path = r"C:\path\to\your\merged_emails"
    merge_text_files(input_folder, output_base_path)