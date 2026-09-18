// vybe-test: c/doom_startup/atexit_function_pointer_list

#include <assert.h>
#include <stdlib.h>

typedef void (*atexit_func_t)(void);

typedef struct atexit_listentry_s atexit_listentry_t;

struct atexit_listentry_s
{
    atexit_func_t func;
    int run_on_error;
    atexit_listentry_t *next;
};

static atexit_listentry_t *exit_funcs = NULL;
static int calls = 0;

static void mark_exit(void)
{
    calls += 1;
}

static void register_exit(atexit_func_t func, int run_on_error)
{
    atexit_listentry_t *entry;

    entry = malloc(sizeof(*entry));
    entry->func = func;
    entry->run_on_error = run_on_error;
    entry->next = exit_funcs;
    exit_funcs = entry;
}

int main(void)
{
    register_exit(mark_exit, 0);

    assert(exit_funcs != NULL);
    assert(exit_funcs->run_on_error == 0);
    assert(exit_funcs->next == NULL);

    exit_funcs->func();
    assert(calls == 1);
    return 0;
}
