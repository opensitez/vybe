// vybe-test: c/file_io/fwrite_fread_binary
// origin: languages/c/tests/c/test_file_io.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {
    int data[3] = {10, 20, 30};
    FILE *f = fopen("/tmp/vybe_test_fwrite.bin", "wb");
    fwrite(data, sizeof(int), 3, f);
    fclose(f);
    int readback[3] = {0, 0, 0};
    f = fopen("/tmp/vybe_test_fwrite.bin", "rb");
    fread(readback, sizeof(int), 3, f);
    fclose(f);
    printf("%d %d %d\n", readback[0], readback[1], readback[2]);
    return 0;
}

