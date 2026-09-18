void *malloc(unsigned long size);
int strcmp(const char *a, const char *b);

static char *exedir;

static char *make_dir(void)
{
    char *p = malloc(3);
    p[0] = '.';
    p[1] = '/';
    p[2] = '\0';
    return p;
}

int main(void)
{
    exedir = make_dir();
    return strcmp(exedir, "./");
}
