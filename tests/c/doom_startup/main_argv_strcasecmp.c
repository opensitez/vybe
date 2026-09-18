int strcasecmp(const char *a, const char *b);

static int seen(int argc, char **argv)
{
    int i;

    for (i = 1; i < argc && argv[i]; i++)
    {
        if (!strcasecmp("-missing", argv[i]))
        {
            return 1;
        }
    }

    return 0;
}

int main(int argc, char **argv)
{
    return seen(argc, argv);
}
