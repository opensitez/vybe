// vybe-test: c/data_structures/stack_push_pop
// origin: languages/c/tests/c/test_data_structures.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#define STACK_MAX 10
typedef struct { int data[STACK_MAX]; int top; } Stack;
void push(Stack *s, int v) { s->data[s->top++] = v; }
int pop(Stack *s) { return s->data[--s->top]; }
int empty(Stack *s) { return s->top == 0; }
int main() {
const char *__w[] = {"3 2 1\n"};
int __n = 1, __i = 0;

Stack s = {{0}, 0};
push(&s, 1); push(&s, 2); push(&s, 3);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", pop(&s), pop(&s), pop(&s));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

