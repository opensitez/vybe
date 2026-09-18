// vybe-test: c/doom_startup/derived_struct_pointer_from_member_alias
#include <stdio.h>
#include <stdlib.h>

typedef unsigned char byte;

typedef struct node_s
{
    int size;
    struct node_s *next;
} node_t;

typedef struct
{
    node_t *rover;
} zone_t;

int main(void)
{
    zone_t *zone;
    node_t *block;
    node_t *base;
    node_t *derived;

    zone = (zone_t *) malloc(128);
    block = (node_t *) ((byte *) zone + sizeof(zone_t));
    block->size = 77;
    zone->rover = block;
    base = zone->rover;
    derived = (node_t *) ((byte *) base + 32);
    derived->size = 1234;

    printf("%d %d %d %d\n", base == block, base->size, derived->size, block->size);
    return 0;
}
