use super::helpers::*;

// Higher-level data structure patterns using C primitives
macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<stdlib.h>", "<string.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    stack_push_pop => {
        declarations: r#"
#define STACK_MAX 10
typedef struct { int data[STACK_MAX]; int top; } Stack;
void push(Stack *s, int v) { s->data[s->top++] = v; }
int pop(Stack *s) { return s->data[--s->top]; }
int empty(Stack *s) { return s->top == 0; }
"#,
        body: r#"
Stack s = {{0}, 0};
push(&s, 1); push(&s, 2); push(&s, 3);
printf("%d %d %d\n", pop(&s), pop(&s), pop(&s));
return 0;
"#,
        expect: ["3 2 1"]
    },
    queue_enqueue_dequeue => {
        declarations: r#"
#define Q_MAX 10
typedef struct { int data[Q_MAX]; int head; int tail; } Queue;
void enqueue(Queue *q, int v) { q->data[q->tail++ % Q_MAX] = v; }
int dequeue(Queue *q) { return q->data[q->head++ % Q_MAX]; }
"#,
        body: r#"
Queue q = {{0}, 0, 0};
enqueue(&q, 10); enqueue(&q, 20); enqueue(&q, 30);
printf("%d %d %d\n", dequeue(&q), dequeue(&q), dequeue(&q));
return 0;
"#,
        expect: ["10 20 30"]
    },
    hashmap_simple_open_address => {
        declarations: r#"
#define HT_SIZE 16
typedef struct { int key; int val; int used; } Entry;
typedef struct { Entry entries[HT_SIZE]; } HashTable;
void ht_set(HashTable *ht, int key, int val) {
    int idx = (key * 2654435761u) % HT_SIZE;
    ht->entries[idx].key = key; ht->entries[idx].val = val; ht->entries[idx].used = 1;
}
int ht_get(HashTable *ht, int key, int *found) {
    int idx = (key * 2654435761u) % HT_SIZE;
    if (ht->entries[idx].used && ht->entries[idx].key == key) { *found = 1; return ht->entries[idx].val; }
    *found = 0; return 0;
}
"#,
        body: r#"
HashTable ht = {{{0}}};
ht_set(&ht, 5, 100);
ht_set(&ht, 7, 200);
int found;
printf("%d %d\n", ht_get(&ht, 5, &found), ht_get(&ht, 7, &found));
return 0;
"#,
        expect: ["100 200"]
    },
    dynamic_array_append => {
        declarations: r#"
typedef struct { int *data; int len; int cap; } Vec;
void vec_push(Vec *v, int x) {
    if (v->len >= v->cap) {
        v->cap = v->cap ? v->cap * 2 : 4;
        v->data = (int*)realloc(v->data, v->cap * sizeof(int));
    }
    v->data[v->len++] = x;
}
"#,
        body: r#"
Vec v = {NULL, 0, 0};
for (int i = 0; i < 5; i++) vec_push(&v, i * 10);
for (int i = 0; i < v.len; i++) printf("%d\n", v.data[i]);
free(v.data);
return 0;
"#,
        expect: ["0", "10", "20", "30", "40"]
    }
}
