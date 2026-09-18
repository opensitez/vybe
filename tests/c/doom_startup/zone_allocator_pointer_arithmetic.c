// vybe-test: c/doom_startup/zone_allocator_pointer_arithmetic
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
    int total = 1024;

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

    printf("%d %d %d %d\n",
           (int) sizeof(memzone_t),
           block->size,
           block == zone->rover,
           block->next == &zone->blocklist);
    return 0;
}
