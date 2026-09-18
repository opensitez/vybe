// vybe-test: c/doom_startup/format_argument_scanner

#include <assert.h>

typedef enum
{
    FORMAT_ARG_INVALID,
    FORMAT_ARG_INT,
    FORMAT_ARG_STRING
} format_arg_t;

static format_arg_t FormatArgumentType(char c)
{
    switch (c)
    {
        case 'd':
            return FORMAT_ARG_INT;

        case 's':
            return FORMAT_ARG_STRING;

        default:
            return FORMAT_ARG_INVALID;
    }
}

static format_arg_t NextFormatArgument(const char **str)
{
    format_arg_t argtype;

    while (**str != '\0')
    {
        if (**str == '%')
        {
            ++*str;

            if (**str != '%')
            {
                break;
            }
        }

        ++*str;
    }

    while (**str != '\0')
    {
        argtype = FormatArgumentType(**str);

        if (argtype != FORMAT_ARG_INVALID)
        {
            ++*str;
            return argtype;
        }

        ++*str;
    }

    *str = 0;

    return FORMAT_ARG_INVALID;
}

int main(void)
{
    const char *s = "%% literal %04d then %s";

    assert(NextFormatArgument(&s) == FORMAT_ARG_INT);
    assert(*s == ' ');
    assert(NextFormatArgument(&s) == FORMAT_ARG_STRING);
    assert(*s == '\0');
    assert(NextFormatArgument(&s) == FORMAT_ARG_INVALID);
    assert(s == 0);

    return 0;
}
