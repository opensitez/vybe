// vybe-test: c/doom_startup/zone_allocator_split_block
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
    memblock_t *newblock;
    int total = 1024;
    int size = 104;
    int extra;

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

    base = zone->rover;
    extra = base->size - size;
    newblock = (memblock_t *) ((byte *) base + size);
    newblock->size = extra;
    newblock->tag = 0;
    newblock->user = NULL;
    newblock->prev = base;
    newblock->next = base->next;
    newblock->next->prev = newblock;
    base->next = newblock;
    base->size = size;
    zone->rover = base->next;

    printf("%d %d %d %d %d %d %d %d\n",
           base->size,
           extra,
           newblock->size,
           zone->rover == newblock,
           newblock->prev == base,
           newblock->next == &zone->blocklist,
           zone->blocklist.prev == newblock,
           base->next == newblock);

    return 0;
}
