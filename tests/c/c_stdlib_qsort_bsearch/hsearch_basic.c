// vybe-test: c/c_stdlib_qsort_bsearch/hsearch_basic
// origin: languages/c/tests/c/test_c_stdlib_qsort_bsearch.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _GNU_SOURCE
#include <search.h>
#include <stdlib.h>
int main() {const char *__w[] = {"42"};
int __n = 1, __i = 0;
 hcreate(10); ENTRY e, *ep; e.key = "foo"; e.data = (void*)42; hsearch(e, ENTER); e.key = "foo"; ep = hsearch(e, FIND); { char __t[512]; snprintf(__t, sizeof(__t), "%ld", (long)ep->data);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } hdestroy(); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

