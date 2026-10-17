#include <stdio.h>
#include <stdlib.h>
static int payload = 42;
static void status_handler(int status, void *arg) {
    printf("%d:%d", status, *(int *)arg);
}
static void relay(int status, void *argument) {
    status_handler(status, argument);
}
int main(void) {
    int local = 66;
    int *linear = (int *)malloc(sizeof(int));
    *linear = 77;
    int *argument = &payload;
    status_handler(7, argument);
    payload = 55;
    printf(",");
    status_handler(7, &payload);
    printf(",");
    relay(7, &local);
    printf(",");
    relay(7, linear);
    return 0;
}
