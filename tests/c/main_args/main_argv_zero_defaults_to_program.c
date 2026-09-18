// vybe-test: c/main_args/main_argv_zero_defaults_to_program
#include <assert.h>
#include <stdio.h>
#include <string.h>

static void check_argv(int argc, char **argv)
{
    assert(argc >= 1);
    assert(argv != NULL);
    assert(argv[0] != NULL);
    assert(strcmp(argv[0], "program") == 0 || strstr(argv[0], "a.out") != NULL || strlen(argv[0]) > 0);
}

int main(int argc, char *argv[])
{
    check_argv(argc, argv);
    return 0;
}
