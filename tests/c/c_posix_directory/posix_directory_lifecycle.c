// vybe-test: c/c_posix_directory/posix_directory_lifecycle
// origin: languages/c/tests/c/test_c_posix_directory.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#define _POSIX_C_SOURCE 200809L
#include <sys/stat.h>
#include <dirent.h>
#include <unistd.h>
#include <string.h>
#include <fcntl.h>

int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;

    const char *dir_name = "test_dir_lifecycle";
    mkdir(dir_name, 0755);
    
    // Create a file inside
    char filepath[256];
    snprintf(filepath, sizeof(filepath), "%s/file1.txt", dir_name);
    int fd = open(filepath, O_CREAT | O_WRONLY, 0644);
    close(fd);
    
    DIR *d = opendir(dir_name);
    if (!d) return 1;
    
    int found_file = 0;
    struct dirent *dir;
    while ((dir = readdir(d)) != NULL) {
        if (strcmp(dir->d_name, "file1.txt") == 0) {
            found_file = 1;
        }
    }
    closedir(d);
    
    unlink(filepath);
    rmdir(dir_name);
    
    { char __t[512]; snprintf(__t, sizeof(__t), "%d", found_file);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

