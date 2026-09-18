#ifndef ON_EXIT_COMPAT_H
#define ON_EXIT_COMPAT_H

#include <stdlib.h>

#if defined(__APPLE__) || !defined(__GLIBC__)
typedef void (*__on_exit_fn)(int, void *);

typedef struct {
    __on_exit_fn fn;
    void *arg;
} __on_exit_slot;

static __on_exit_slot __on_exit_slots[32];
static int __on_exit_count = 0;
static int __exit_code = 0;

static void __on_exit_trampoline(void) {
    if (__on_exit_count > 0) {
        __on_exit_count--;
        __on_exit_slots[__on_exit_count].fn(__exit_code, __on_exit_slots[__on_exit_count].arg);
    }
}

static inline int on_exit(__on_exit_fn fn, void *arg) {
    if (__on_exit_count < 32) {
        __on_exit_slots[__on_exit_count].fn = fn;
        __on_exit_slots[__on_exit_count].arg = arg;
        __on_exit_count++;
        return atexit(__on_exit_trampoline);
    }
    return -1;
}

static inline void __custom_exit(int status) {
    __exit_code = status;
    exit(status);
}
#define exit(s) __custom_exit(s)

#endif

#endif
