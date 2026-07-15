use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn mblen_ascii() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int len = mblen(\"a\", 1); printf(\"%d\", len); return 0; }"), vec!["1"]); }
#[test] fn mblen_null() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int len = mblen(NULL, 0); printf(\"%d\", len == 0); return 0; }"), vec!["1"]); } // query state dependent encodings
#[test] fn mblen_empty_string() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int len = mblen(\"\", 1); printf(\"%d\", len); return 0; }"), vec!["0"]); }
#[test] fn mblen_invalid() { assert_eq!(run_c("#include <stdlib.h>\nint main() { char inv[] = {(char)0xff}; int len = mblen(inv, 1); printf(\"%d\", len); return 0; }"), vec!["-1"]); } // assuming standard utf8/c locale behavior
#[test] fn mbtowc_ascii() { assert_eq!(run_c("#include <stdlib.h>\nint main() { wchar_t wc; int len = mbtowc(&wc, \"A\", 1); printf(\"%d %d\", len, (int)wc); return 0; }"), vec!["1 65"]); }
#[test] fn mbtowc_null_ptr() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int len = mbtowc(NULL, \"A\", 1); printf(\"%d\", len); return 0; }"), vec!["1"]); } // still returns length
#[test] fn mbtowc_null_string() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int len = mbtowc(NULL, NULL, 0); printf(\"%d\", len == 0); return 0; }"), vec!["1"]); } // state query
#[test] fn wctomb_ascii() { assert_eq!(run_c("#include <stdlib.h>\nint main() { char s[MB_CUR_MAX]; int len = wctomb(s, L'B'); printf(\"%d %c\", len, s[0]); return 0; }"), vec!["1 B"]); }
#[test] fn wctomb_null_ptr() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int len = wctomb(NULL, L'A'); printf(\"%d\", len == 0); return 0; }"), vec!["1"]); } // state query
#[test] fn mbstowcs_ascii() { assert_eq!(run_c("#include <stdlib.h>\nint main() { wchar_t ws[10]; size_t len = mbstowcs(ws, \"hello\", 10); printf(\"%d %d\", (int)len, (int)ws[0]); return 0; }"), vec!["5 104"]); }
#[test] fn mbstowcs_null_dst() { assert_eq!(run_c("#include <stdlib.h>\nint main() { size_t len = mbstowcs(NULL, \"hello\", 0); printf(\"%d\", (int)len); return 0; }"), vec!["5"]); } // query length needed
#[test] fn wcstombs_ascii() { assert_eq!(run_c("#include <stdlib.h>\nint main() { char s[10]; size_t len = wcstombs(s, L\"hello\", 10); printf(\"%d %c\", (int)len, s[0]); return 0; }"), vec!["5 h"]); }
#[test] fn wcstombs_null_dst() { assert_eq!(run_c("#include <stdlib.h>\nint main() { size_t len = wcstombs(NULL, L\"hello\", 0); printf(\"%d\", (int)len); return 0; }"), vec!["5"]); } // query length needed
#[test] fn wcstombs_truncation() { assert_eq!(run_c("#include <stdlib.h>\nint main() { char s[4]; size_t len = wcstombs(s, L\"hello\", 3); printf(\"%d %c%c%c\", (int)len, s[0], s[1], s[2]); return 0; }"), vec!["3 hel"]); }
#[test] fn mbstowcs_truncation() { assert_eq!(run_c("#include <stdlib.h>\nint main() { wchar_t ws[4]; size_t len = mbstowcs(ws, \"hello\", 3); printf(\"%d %d\", (int)len, (int)ws[2]); return 0; }"), vec!["3 108"]); } // 108 = 'l'
#[test] fn mblen_mb_cur_max() { assert_eq!(run_c("#include <stdlib.h>\nint main() { printf(\"%d\", MB_CUR_MAX >= 1); return 0; }"), vec!["1"]); }
#[test] fn wctomb_invalid_char() { assert_eq!(run_c("#include <stdlib.h>\nint main() { char s[MB_CUR_MAX]; /* wctomb doesn't validate much in C locale, but it shouldn't crash */ int len = wctomb(s, (wchar_t)-1); printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn mbtowc_empty() { assert_eq!(run_c("#include <stdlib.h>\nint main() { wchar_t wc; int len = mbtowc(&wc, \"\", 1); printf(\"%d %d\", len, (int)wc); return 0; }"), vec!["0 0"]); }
#[test] fn mbstowcs_empty() { assert_eq!(run_c("#include <stdlib.h>\nint main() { wchar_t ws[10]; size_t len = mbstowcs(ws, \"\", 10); printf(\"%d %d\", (int)len, (int)ws[0]); return 0; }"), vec!["0 0"]); }
#[test] fn wcstombs_empty() { assert_eq!(run_c("#include <stdlib.h>\nint main() { char s[10]; size_t len = wcstombs(s, L\"\", 10); printf(\"%d %d\", (int)len, s[0]); return 0; }"), vec!["0 0"]); }
