// vybe-test: c/doom_startup/char_pointer_cell_preincrement

#include <assert.h>

static int count_until_end(const char **p)
{
    int n = 0;

    while (**p != '\0')
    {
        ++*p;
        ++n;
    }

    return n;
}

int main(void)
{
    const char *s = "abc";

    assert(*s == 'a');
    assert(count_until_end(&s) == 3);
    assert(*s == '\0');

    return 0;
}
