typedef unsigned char byte;

typedef struct {
    byte data[8];
    int pos;
} packet_t;

static int isprint(int c) {
    return c >= 32 && c <= 126;
}

char *read_string(packet_t *packet) {
    return (char *) &packet->data[packet->pos];
}

char *safe_string(packet_t *packet) {
    char *r;
    char *w;
    char *result;

    result = read_string(packet);
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
    packet_t packet = {{'A', 1, 'B', '\n', 'C', 0, 'Z', 0}, 0};
    char *out = safe_string(&packet);
    return out[0] == 'A' && out[1] == 'B' && out[2] == '\n' && out[3] == 'C' && out[4] == 0 ? 0 : 1;
}
