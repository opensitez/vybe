#include "geometry.h"
#include "counter.h"
int rect_area(struct Rect r) { count_call(); return r.w * r.h; }
int rect_perimeter(struct Rect r) { count_call(); return 2 * (r.w + r.h); }
