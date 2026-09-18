void *malloc(unsigned long size);
char *strdup(const char *s);
int strcasecmp(const char *a, const char *b);

static int myargc;
static char **myargv;

static char *duplicate(const char *s)
{
    return strdup(s);
}

static int check(const char *name)
{
    int i;

    for (i = 1; i < myargc && myargv[i]; i++)
    {
        if (!strcasecmp(name, myargv[i]))
        {
            return i;
        }
    }

    return 0;
}

int main(int argc, char **argv)
{
    int i;

    myargc = argc;
    myargv = malloc(argc * sizeof(char *));

    for (i = 0; i < argc; i++)
    {
        myargv[i] = duplicate(argv[i]);
    }

    return check("-missing");
}
