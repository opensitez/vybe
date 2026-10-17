struct node { int value; };
int __c_static_iife_0 = 17;
static struct node original = {7};
static int calls;

static struct node *advance(void) {
    ++calls;
    return &original;
}

int main(void) {
    int __c_static_iife_1 = 23;
    if (advance() != advance()) return 1;
    if (calls != 2) return 2;
    if (__c_static_iife_0 != 17) return 3;
    if (__c_static_iife_1 != 23) return 4;
    return 0;
}
