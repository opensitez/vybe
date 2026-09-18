void *malloc(unsigned long size);
int strcasecmp(const char *a, const char *b);

static int myargc;
static char **myargv;

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
        myargv[i] = argv[i];
    }

    return check("-missing");
}
