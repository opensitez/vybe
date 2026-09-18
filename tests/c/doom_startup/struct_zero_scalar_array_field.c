// vybe-test: c/doom_startup/struct_zero_scalar_array_field

#include <assert.h>

typedef struct
{
    int values[256];
} bucket_t;

int main(void)
{
    bucket_t bucket = {0};

    assert(bucket.values[0] == 0);
    assert(bucket.values[255] == 0);

    bucket.values[17] = 42;
    assert(bucket.values[17] == 42);
    assert(bucket.values[18] == 0);

    return 0;
}
