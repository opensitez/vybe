typedef unsigned char byte;

static int isprint(int c) {
    return c >= 32 && c <= 126;
}

char *safe_string(char *result) {
    char *r;
    char *w;

    if (result == 0) {
        return 0;
    }

    w = result;
    for (r = result; *r != '\0'; ++r) {
        if (isprint(*r) || *r == '\n') {
            *w = *r;
            ++w;
        }
    }
    *w = '\0';

    return result;
}

int main(void) {
    char text[6] = { 'A', 1, 'B', '\n', 'C', 0 };
    char *out = safe_string(text);
    return out[0] == 'A' && out[1] == 'B' && out[2] == '\n' && out[3] == 'C' && out[4] == 0 ? 0 : 1;
}
