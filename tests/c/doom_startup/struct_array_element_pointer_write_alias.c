// vybe-test: c/doom_startup/struct_array_element_pointer_write_alias
#include <stdio.h>
#include <stdlib.h>

typedef struct lumpinfo_s lumpinfo_t;

struct lumpinfo_s
{
    char name[8];
    int position;
    int size;
    void *cache;
    int next;
};

int main(void)
{
    lumpinfo_t *filelumps = calloc(2, sizeof(lumpinfo_t));
    lumpinfo_t **lumpinfo = calloc(2, sizeof(lumpinfo_t *));
    lumpinfo_t *lump_p;

    lump_p = &filelumps[1];
    lump_p->position = 12;
    lump_p->size = 2920;
    lump_p->name[0] = 'T';
    lump_p->name[1] = 'H';
    lump_p->name[2] = 'I';
    lump_p->name[3] = 'N';
    lump_p->name[4] = 'G';
    lump_p->name[5] = 'S';
    lumpinfo[1] = lump_p;

    printf("%d %d %.8s\n", filelumps[1].position, filelumps[1].size, filelumps[1].name);
    printf("%d %d %.8s\n", lumpinfo[1]->position, lumpinfo[1]->size, lumpinfo[1]->name);
    return 0;
}
