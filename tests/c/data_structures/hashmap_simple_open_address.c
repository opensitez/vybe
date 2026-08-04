// vybe-test: c/data_structures/hashmap_simple_open_address
// origin: languages/c/tests/c/test_data_structures.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

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
int main() {
const char *__w[] = {"100 200\n"};
int __n = 1, __i = 0;

HashTable ht = {{{0}}};
ht_set(&ht, 5, 100);
ht_set(&ht, 7, 200);
int found;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", ht_get(&ht, 5, &found), ht_get(&ht, 7, &found));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

