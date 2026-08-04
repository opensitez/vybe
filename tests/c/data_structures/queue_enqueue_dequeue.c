// vybe-test: c/data_structures/queue_enqueue_dequeue
// origin: languages/c/tests/c/test_data_structures.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#define Q_MAX 10
typedef struct { int data[Q_MAX]; int head; int tail; } Queue;
void enqueue(Queue *q, int v) { q->data[q->tail++ % Q_MAX] = v; }
int dequeue(Queue *q) { return q->data[q->head++ % Q_MAX]; }
int main() {
const char *__w[] = {"10 20 30\n"};
int __n = 1, __i = 0;

Queue q = {{0}, 0, 0};
enqueue(&q, 10); enqueue(&q, 20); enqueue(&q, 30);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", dequeue(&q), dequeue(&q), dequeue(&q));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

