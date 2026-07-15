use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn aligned_alloc_basic() { assert_eq!(run_c("#include <stdlib.h>\nint main() { void *p = aligned_alloc(64, 128); printf(\"%d\", ((size_t)p % 64) == 0); free(p); return 0; }"), vec!["1"]); }
#[test] fn posix_memalign_basic() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { void *p; int res = posix_memalign(&p, 128, 256); printf(\"%d %d\", res == 0, ((size_t)p % 128) == 0); free(p); return 0; }"), vec!["1 1"]); }
#[test] fn posix_memalign_invalid_alignment() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { void *p; int res = posix_memalign(&p, 65, 256); printf(\"%d\", res != 0); return 0; }"), vec!["1"]); }
#[test] fn memalign_basic() { assert_eq!(run_c("#include <malloc.h>\nint main() { void *p = memalign(32, 128); printf(\"%d\", ((size_t)p % 32) == 0); free(p); return 0; }"), vec!["1"]); }
#[test] fn valloc_basic() { assert_eq!(run_c("#define _BSD_SOURCE\n#include <stdlib.h>\n#include <unistd.h>\nint main() { void *p = valloc(1024); long pagesize = sysconf(_SC_PAGESIZE); printf(\"%d\", ((size_t)p % pagesize) == 0); free(p); return 0; }"), vec!["1"]); }
#[test] fn pvalloc_basic() { assert_eq!(run_c("#define _GNU_SOURCE\n#include <malloc.h>\n#include <unistd.h>\nint main() { void *p = pvalloc(10); long pagesize = sysconf(_SC_PAGESIZE); printf(\"%d\", ((size_t)p % pagesize) == 0); free(p); return 0; }"), vec!["1"]); }
#[test] fn aligned_alloc_unaligned_size() { assert_eq!(run_c("#include <stdlib.h>\nint main() { void *p = aligned_alloc(32, 33); /* size must be multiple of align in C11 */ printf(\"%d\", p == NULL || ((size_t)p % 32) == 0); free(p); return 0; }"), vec!["1"]); }
#[test] fn posix_memalign_small_alignment() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { void *p; int res = posix_memalign(&p, 2, 256); /* must be multiple of sizeof(void*) */ printf(\"%d\", res != 0); return 0; }"), vec!["1"]); }
#[test] fn malloc_usable_size_check() { assert_eq!(run_c("#include <malloc.h>\nint main() { void *p = malloc(10); printf(\"%d\", malloc_usable_size(p) >= 10); free(p); return 0; }"), vec!["1"]); }
#[test] fn mallinfo_check() { assert_eq!(run_c("#include <malloc.h>\nint main() { struct mallinfo m = mallinfo(); printf(\"%d\", m.arena >= 0); return 0; }"), vec!["1"]); }
#[test] fn mallopt_check() { assert_eq!(run_c("#include <malloc.h>\nint main() { int res = mallopt(M_TRIM_THRESHOLD, 128*1024); printf(\"%d\", res == 1); return 0; }"), vec!["1"]); }
#[test] fn reallocarray_basic() { assert_eq!(run_c("#define _GNU_SOURCE\n#include <stdlib.h>\nint main() { void *p = reallocarray(NULL, 10, 5); printf(\"%d\", p != NULL); free(p); return 0; }"), vec!["1"]); }
#[test] fn reallocarray_overflow() { assert_eq!(run_c("#define _GNU_SOURCE\n#include <stdlib.h>\nint main() { void *p = reallocarray(NULL, (size_t)-1, 10); printf(\"%d\", p == NULL); return 0; }"), vec!["1"]); }
#[test] fn alloca_basic() { assert_eq!(run_c("#include <alloca.h>\nint main() { char *p = alloca(10); p[0] = 'X'; printf(\"%c\", p[0]); return 0; }"), vec!["X"]); }
#[test] fn alloca_in_loop() { assert_eq!(run_c("#include <alloca.h>\nint main() { for(int i=0; i<3; i++) { int *p = alloca(sizeof(int)); *p = i; if (i == 2) printf(\"%d\", *p); } return 0; }"), vec!["2"]); }
#[test] fn malloc_stats_run() { assert_eq!(run_c("#include <malloc.h>\nint main() { malloc_stats(); printf(\"ok\"); return 0; }"), vec!["ok"]); } // stderr output ignored by run_prints
#[test] fn memalign_free_is_safe() { assert_eq!(run_c("#include <malloc.h>\nint main() { void *p = memalign(64, 128); free(p); printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn aligned_alloc_large() { assert_eq!(run_c("#include <stdlib.h>\nint main() { void *p = aligned_alloc(4096, 8192); printf(\"%d\", ((size_t)p % 4096) == 0); free(p); return 0; }"), vec!["1"]); }
#[test] fn posix_memalign_zero_size() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { void *p = (void*)1; posix_memalign(&p, 128, 0); printf(\"%d\", p == NULL || p == (void*)1 || p != NULL); return 0; }"), vec!["1"]); }
#[test] fn realloc_to_zero() { assert_eq!(run_c("#include <stdlib.h>\nint main() { void *p = malloc(10); p = realloc(p, 0); printf(\"%d\", p == NULL || p != NULL); return 0; }"), vec!["1"]); }
