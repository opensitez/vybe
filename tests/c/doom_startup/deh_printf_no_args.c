// vybe-test: c/doom_startup/deh_printf_no_args

#include <stdarg.h>
#include <stdio.h>

static const char *identity(const char *s)
{
    return s;
}

static void deh_printf(const char *fmt, ...)
{
    va_list args;
    const char *repl;

    repl = identity(fmt);
    va_start(args, fmt);
    vprintf(repl, args);
    va_end(args);
}

int main(void)
{
    puts("banner");
    deh_printf("Z_Init: Init zone memory allocation daemon.\n");
    puts("after");
    return 0;
}
