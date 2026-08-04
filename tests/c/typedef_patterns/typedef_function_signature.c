// vybe-test: c/typedef_patterns/typedef_function_signature
// origin: languages/c/tests/c/test_typedef_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef void (*Handler)(int);
void on_event(int code) { printf("%d\n", code); }
int main() {
Handler h = on_event;
h(42);
return 0;
}

