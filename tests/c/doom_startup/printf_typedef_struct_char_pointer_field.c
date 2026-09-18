typedef __builtin_va_list va_list;

typedef struct
{
    int *defaults;
    int numdefaults;
    const char *filename;
} default_collection_t;

static default_collection_t doom_defaults = { 0, 0, 0 };
static void *collection_addr = 0;

int snprintf(char *buf, unsigned long size, const char *fmt, ...);
int strcmp(const char *a, const char *b);

int main(void)
{
    char out[32];

    collection_addr = &doom_defaults;
    doom_defaults.filename = "default.cfg";
    snprintf(out, sizeof(out), "%s", doom_defaults.filename);
    return strcmp(out, "default.cfg");
}
