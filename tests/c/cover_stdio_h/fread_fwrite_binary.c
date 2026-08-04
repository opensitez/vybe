// vybe-test: c/cover_stdio_h/fread_fwrite_binary
// origin: languages/c/tests/c/test_cover_stdio_h.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
FILE *f=fopen("/tmp/vybe_c_bin.dat","wb"); int v=7; fwrite(&v,sizeof(v),1,f); fclose(f); f=fopen("/tmp/vybe_c_bin.dat","rb"); int o=0; fread(&o,sizeof(o),1,f); fclose(f); printf("%d\n", o); return 0;
}

