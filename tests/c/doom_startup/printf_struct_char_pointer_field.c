typedef __builtin_va_list va_list;

struct config {
    char *filename;
};

int snprintf(char *buf, unsigned long size, const char *fmt, ...);
int strcmp(const char *a, const char *b);

int main(void)
{
    struct config cfg;
    char out[32];

    cfg.filename = "default.cfg";
    snprintf(out, sizeof(out), "%s", cfg.filename);
    return strcmp(out, "default.cfg");
}
