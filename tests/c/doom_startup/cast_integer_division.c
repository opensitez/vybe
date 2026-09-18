// vybe-test: c/doom_startup/cast_integer_division
#include <stdio.h>

int main(void)
{
    printf("%d %d %d\n", 24 / 8, (int) (24 / 8), (int) 3);
    return 0;
}
