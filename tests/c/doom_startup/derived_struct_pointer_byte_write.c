// vybe-test: c/doom_startup/derived_struct_pointer_byte_write
#include <stdio.h>
#include <stdlib.h>

typedef unsigned char byte;

typedef struct node_s
{
    int size;
    struct node_s *next;
} node_t;

int main(void)
{
    node_t *base;
    node_t *derived;
    byte *raw;

    base = (node_t *) malloc(128);
    base->size = 11;
    derived = (node_t *) ((byte *) base + 32);
    derived->size = 0x01020304;
    raw = (byte *) derived;

    printf("%d %d %d %d %d\n",
           derived->size,
           raw[0],
           raw[1],
           raw[2],
           raw[3]);
    return 0;
}
