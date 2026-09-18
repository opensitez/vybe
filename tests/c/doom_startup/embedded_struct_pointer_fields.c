// vybe-test: c/doom_startup/embedded_struct_pointer_fields
#include <stdio.h>
#include <stdlib.h>

typedef unsigned char byte;

typedef struct node_s
{
    int size;
    struct node_s *next;
    struct node_s *prev;
} node_t;

typedef struct
{
    int size;
    node_t sentinel;
    node_t *rover;
} zone_t;

int main(void)
{
    zone_t *zone;
    zone = (zone_t *) malloc(256);
    node_t *block = (node_t *) ((byte *) zone + sizeof(zone_t));
    node_t *split = (node_t *) ((byte *) block + 40);

    zone->sentinel.next = block;
    zone->sentinel.prev = block;
    block->next = &zone->sentinel;
    block->prev = &zone->sentinel;
    split->next = block->next;
    split->next->prev = split;

    printf("%d %d %d %d %d\n",
           zone->sentinel.next == block,
           zone->sentinel.prev == block,
           block->next == &zone->sentinel,
           split->next->prev == split,
           zone->sentinel.prev == split);
    return 0;
}
