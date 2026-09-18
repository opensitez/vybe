typedef __builtin_va_list va_list;

int vsnprintf(char *buf, unsigned long size, const char *fmt, va_list ap);
int strcmp(const char *a, const char *b);
void va_start(va_list ap, ...);
void va_end(va_list ap);

static int M_vsnprintf(char *buf, unsigned long buf_len, const char *s, va_list args)
{
    int result;

    if (buf_len < 1)
    {
        return 0;
    }

    result = vsnprintf(buf, buf_len, s, args);

    if (result < 0 || result >= buf_len)
    {
        buf[buf_len - 1] = '\0';
        result = buf_len - 1;
    }

    return result;
}

static int M_snprintf(char *buf, unsigned long buf_len, const char *s, ...)
{
    va_list args;
    int result;

    va_start(args, s);
    result = M_vsnprintf(buf, buf_len, s, args);
    va_end(args);

    return result;
}

int main(void)
{
    char name[32];

    M_snprintf(name, sizeof(name), "joystick_physical_button%i", 3);

    return strcmp(name, "joystick_physical_button3");
}
