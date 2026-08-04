// vybe-test: c/file_io/fprintf_writes_formatted
// origin: languages/c/tests/c/test_file_io.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {
    FILE *f = fopen("/tmp/vybe_test_fprintf.txt", "w");
    fprintf(f, "%d %s\n", 42, "test");
    fclose(f);
    f = fopen("/tmp/vybe_test_fprintf.txt", "r");
    char buf[50];
    fgets(buf, sizeof(buf), f);
    fclose(f);
    printf("%s", buf);
    return 0;
}

