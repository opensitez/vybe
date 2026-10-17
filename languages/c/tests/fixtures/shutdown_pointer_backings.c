#include <stdio.h>
#include <stdlib.h>
struct State { int value; };
static struct State state;
static int payload = 42;
static void scalar_handler(int status, void *argument) {
    if (status != 0) _Exit(99);
    printf("%d,", *(int *)argument);
}
static void record_handler(int status, void *argument) {
    struct State *record = (struct State *)argument;
    if (status != 0) _Exit(99);
    printf("%d,", record->value);
}
int main(void) {
    int *linear = (int *)malloc(sizeof(int));
    *linear = 77;
    on_exit(scalar_handler, linear);
    on_exit(record_handler, &state);
    on_exit(scalar_handler, &payload);
    payload = 66;
    state.value = 55;
    *linear = 88;
    return 0;
}
