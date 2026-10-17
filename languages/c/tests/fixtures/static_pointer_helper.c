#include <stdlib.h>

struct node { int value; };
static struct node *first;
static struct node *second;
static int left_calls;
static int right_calls;

static struct node *left(void) { ++left_calls; return first; }
static struct node *right(void) { ++right_calls; return second; }

int main(void)
{
    first = malloc(sizeof(*first));
    second = malloc(sizeof(*second));
    first->value = 7;
    second->value = 7;
    if (left() == right()) return 1;
    if (left_calls != 1 || right_calls != 1) return 2;
    second = first;
    if (left() != right()) return 3;
    if (left_calls != 2 || right_calls != 2) return 4;
    return 0;
}
