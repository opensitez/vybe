// expect: 7

int main(void) {
    int x = 3;
#define ADD1(v) ((v) + 1)
    x = ADD1(x);
#define ADD3(v) ((v) + 3)
    return ADD3(x);
}
