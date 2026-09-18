// vybe-test: c/doom_startup/char_pointer_name_does_not_rewrite_int_compound_assign

#include <assert.h>

static int keep_char_pointer_name_live(void)
{
    char *result = "abc";
    return result[1] == 'b';
}

static int keep_carray_pointer_name_live(void)
{
    int values[2] = {3, 4};
    int *result = values;
    return result[1] == 4;
}

static int joystick_axis_like_numeric_update(void)
{
    int result = 0;
    result -= 32767;
    result += 32767;
    return result;
}

int main(void)
{
    assert(keep_char_pointer_name_live());
    assert(keep_carray_pointer_name_live());
    assert(joystick_axis_like_numeric_update() == 0);
    return 0;
}
