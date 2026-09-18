typedef unsigned char byte;
typedef unsigned long size_t;

typedef struct packet_s {
    byte *data;
    size_t len;
    size_t alloced;
} packet_t;

extern size_t strlen(const char *s);

static void grow(packet_t *packet) {
    packet->alloced = packet->alloced + 16;
}

static void copy(char *dest, const char *src, size_t len) {
    (void) dest;
    (void) src;
    (void) len;
}

void write_string(packet_t *packet, const char *string) {
    byte *p;
    size_t string_size;

    string_size = strlen(string) + 1;

    while (packet->len + string_size > packet->alloced) {
        grow(packet);
    }

    p = packet->data + packet->len;

    copy((char *) p, string, string_size);

    packet->len += string_size;
}

int main(void) {
    byte buf[32];
    packet_t packet = { buf, 0, 32 };
    write_string(&packet, "abc");
    return packet.len == 4 ? 0 : 1;
}
