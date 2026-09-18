typedef __builtin_va_list va_list;

void *malloc(unsigned long size);
unsigned long strlen(const char *s);
char *strncpy(char *dest, const char *src, unsigned long n);
char *strdup(const char *s);
int strcmp(const char *a, const char *b);
void va_start(va_list ap, ...);
void va_end(va_list ap);

static char *make_copy(const char *src)
{
    char *result;
    unsigned long len;

    len = strlen(src) + 1;
    result = malloc(len);
    result[len - 1] = '\0';
    strncpy(result, src, len - 1);

    return result;
}

int main(void)
{
    char *raw;
    char *dup;

    raw = make_copy("./");
    dup = strdup(raw);

    return strcmp(dup, "./");
}
