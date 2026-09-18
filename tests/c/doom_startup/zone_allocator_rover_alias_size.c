// vybe-test: c/doom_startup/zone_allocator_rover_alias_size
#include <stdio.h>
#include <stdlib.h>

typedef unsigned char byte;

typedef struct memblock_s
{
    int size;
    void **user;
    int tag;
    int id;
    struct memblock_s *next;
    struct memblock_s *prev;
} memblock_t;

typedef struct
{
    int size;
    memblock_t blocklist;
    memblock_t *rover;
} memzone_t;

int main(void)
{
    memzone_t *zone;
    memblock_t *block;
    memblock_t *base;

    zone = (memzone_t *) malloc(1024);
    zone->size = 1024;
    block = (memblock_t *) ((byte *) zone + sizeof(memzone_t));
    zone->rover = block;
    block->size = 968;
    base = zone->rover;

    printf("%d %d %d %d\n", (int) sizeof(memzone_t), block->size, base->size, base == block);
    return 0;
}
