// vybe-test: c/doom_startup/zone_self_linked_blocklist

#include <assert.h>
#include <stddef.h>
#include <stdlib.h>
#include <string.h>

typedef unsigned char byte;

typedef struct memblock_s memblock_t;

struct memblock_s
{
    int size;
    void **user;
    int tag;
    memblock_t *next;
    memblock_t *prev;
};

typedef struct
{
    int size;
    memblock_t blocklist;
    memblock_t *rover;
} memzone_t;

static memzone_t *mainzone;

static void *zone_base(int *size)
{
    *size = 4096;
    return malloc(*size);
}

static void zone_init(void)
{
    memblock_t *block;
    int size;

    mainzone = (memzone_t *) zone_base(&size);
    mainzone->size = size;

    mainzone->blocklist.next =
        mainzone->blocklist.prev =
        block = (memblock_t *) ((byte *) mainzone + sizeof(memzone_t));

    mainzone->blocklist.user = (void *) mainzone;
    mainzone->blocklist.tag = 1;
    mainzone->rover = block;

    block->prev = block->next = &mainzone->blocklist;
    block->tag = 0;
    block->size = mainzone->size - sizeof(memzone_t);
}

int main(void)
{
    zone_init();
    assert(mainzone->blocklist.next == mainzone->rover);
    assert(mainzone->blocklist.prev == mainzone->rover);
    assert(mainzone->rover->next == &mainzone->blocklist);
    assert(mainzone->rover->prev == &mainzone->blocklist);
    assert(mainzone->rover->size == 4096 - (int) sizeof(memzone_t));
    return 0;
}
