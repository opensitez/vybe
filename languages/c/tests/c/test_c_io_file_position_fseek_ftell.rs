use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn ftell_start() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_ftell.txt\", \"w+\"); printf(\"%ld\", ftell(f)); fclose(f); return 0; }"), vec!["0"]); }
#[test] fn ftell_after_write() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_ftell_w.txt\", \"w+\"); fputs(\"hello\", f); printf(\"%ld\", ftell(f)); fclose(f); return 0; }"), vec!["5"]); }
#[test] fn fseek_set() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_fseek.txt\", \"w+\"); fputs(\"hello\", f); fseek(f, 1, SEEK_SET); printf(\"%c\", fgetc(f)); fclose(f); return 0; }"), vec!["e"]); }
#[test] fn fseek_cur() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_fseek_cur.txt\", \"w+\"); fputs(\"hello\", f); rewind(f); fgetc(f); fseek(f, 2, SEEK_CUR); printf(\"%c\", fgetc(f)); fclose(f); return 0; }"), vec!["l"]); }
#[test] fn fseek_end() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_fseek_end.txt\", \"w+\"); fputs(\"hello\", f); fseek(f, -2, SEEK_END); printf(\"%c\", fgetc(f)); fclose(f); return 0; }"), vec!["l"]); }
#[test] fn fgetpos_fsetpos() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_fpos.txt\", \"w+\"); fputs(\"hello\", f); fpos_t pos; rewind(f); fgetc(f); fgetpos(f, &pos); fgetc(f); fsetpos(f, &pos); printf(\"%c\", fgetc(f)); fclose(f); return 0; }"), vec!["e"]); }
#[test] fn rewind_clears_eof() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_rewind.txt\", \"w+\"); fputs(\"a\", f); rewind(f); fgetc(f); fgetc(f); printf(\"%d \", feof(f)); rewind(f); printf(\"%d\", feof(f)); fclose(f); return 0; }"), vec!["1 0"]); }
#[test] fn fseek_beyond_end() { assert_eq!(run_c("int main() { FILE *f = fopen(\"test_fseek_beyond.txt\", \"w+\"); fseek(f, 10, SEEK_SET); fputc('A', f); printf(\"%ld\", ftell(f)); fclose(f); return 0; }"), vec!["11"]); }
