struct node { int value; };
static struct node first = {7};
static struct node second = {7};
static int trace;

static struct node *next(int digit) {
    trace = trace * 10 + digit;
    return digit == 1 ? &first : &second;
}

int main(void) {
    if (next(1) == next(2)) return 1;
    if (trace != 12) return 2;
    trace = 0;
    if (next(1) != next(1)) return 3;
    if (trace != 11) return 4;
    if (first.value != 7 || second.value != 7) return 5;
    return 0;
}
