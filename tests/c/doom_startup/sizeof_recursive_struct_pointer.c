// vybe-test: c/doom_startup/sizeof_recursive_struct_pointer
#include <stdio.h>

typedef struct node_s
{
    int size;
    void **user;
    int tag;
    int id;
    struct node_s *next;
    struct node_s *prev;
} node_t;

int main(void)
{
    printf("%d %d %d\n", (int) sizeof(void *), (int) sizeof(node_t *), (int) sizeof(node_t));
    return 0;
}
