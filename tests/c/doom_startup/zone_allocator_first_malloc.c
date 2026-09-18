// vybe-test: c/doom_startup/zone_allocator_first_malloc
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
    memblock_t *rover;
    memblock_t *start;
    int total = 1024;
    int size = 64;

    zone = (memzone_t *) malloc(total);
    zone->size = total;
    zone->blocklist.next =
        zone->blocklist.prev =
        block = (memblock_t *) ((byte *) zone + sizeof(memzone_t));
    zone->blocklist.user = (void *) zone;
    zone->blocklist.tag = 1;
    zone->rover = block;

    block->prev = block->next = &zone->blocklist;
    block->tag = 0;
    block->size = zone->size - sizeof(memzone_t);

    size = (size + sizeof(void *) - 1) & ~(sizeof(void *) - 1);
    size += sizeof(memblock_t);

    base = zone->rover;
    if (base->prev->tag == 0)
    {
        base = base->prev;
    }

    rover = base;
    start = base->prev;

    printf("%d %d %d %d %d %d\n",
           size,
           base->size,
           base->tag,
           base == block,
           start == &zone->blocklist,
           rover == start);

    return 0;
}
