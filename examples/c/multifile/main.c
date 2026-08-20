#include <stdio.h>
#include "geometry.h"
#include "counter.h"

int main(void) {
    struct Rect r = {6, 4};
    printf("area: %d\n", rect_area(r));
    printf("perimeter: %d\n", rect_perimeter(r));
    printf("calls: %d\n", calls_made());
    return 0;
}
