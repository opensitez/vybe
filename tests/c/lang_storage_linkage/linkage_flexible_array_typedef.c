// vybe-test: c/lang_storage_linkage/linkage_flexible_array_typedef
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
// vybe-test-mode: compile
#include <stdlib.h>
typedef struct { int n; char tail[]; } Blob;
int main() {
Blob *b=malloc(sizeof(Blob)+4); free(b); return 0;
}

