#define arrlen(array) (sizeof(array) / sizeof(*array))

typedef struct
{
    const char *name;
    int value;
} entry_t;

static entry_t defaults[] =
{
    { "mouse_sensitivity", 1 },
    { "mouse_acceleration", 2 },
    { "mouse_threshold", 3 },
};

int main(void)
{
    return arrlen(defaults) != 3;
}
