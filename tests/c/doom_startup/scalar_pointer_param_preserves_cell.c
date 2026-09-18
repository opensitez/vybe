// vybe-test: c/doom_startup/scalar_pointer_param_preserves_cell
#include <assert.h>
#include <stdio.h>
#include <string.h>

enum default_type {
    DEFAULT_INT,
    DEFAULT_FLOAT,
};

typedef struct {
    const char *name;
    enum default_type type;
    union {
        int *i;
        float *f;
    } location;
} default_t;

static float mouse_acceleration = 2.0f;
static default_t defaults[] = {
    { "mouse_acceleration", DEFAULT_FLOAT, { 0 } },
};

static default_t *GetDefaultForName(const char *name)
{
    if (strcmp(defaults[0].name, name) == 0)
    {
        return &defaults[0];
    }

    return 0;
}

static void M_BindFloatVariable(const char *name, float *location)
{
    default_t *variable = GetDefaultForName(name);
    assert(variable->type == DEFAULT_FLOAT);
    variable->location.f = location;
}

int main(void)
{
    const char *__w[] = { "7\n" };
    int __n = 1, __i = 0;

    M_BindFloatVariable("mouse_acceleration", &mouse_acceleration);
    *defaults[0].location.f = 7.0f;

    {
        char __t[512];
        snprintf(__t, sizeof(__t), "%.0f\n", mouse_acceleration);
        if (__i >= __n || strcmp(__t, __w[__i]) != 0)
        {
            printf("FAIL at line %d: got [%s]\n", __i, __t);
            assert(0);
        }
        __i++;
    }

    if (__i != __n)
    {
        printf("FAIL: %d line(s), wanted %d\n", __i, __n);
        assert(0);
    }

    return 0;
}
