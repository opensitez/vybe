#define arrlen(array) (sizeof(array) / sizeof(*array))

typedef enum
{
    DEFAULT_INT,
    DEFAULT_FLOAT,
} default_type_t;

typedef struct
{
    const char *name;
    union {
        int *i;
        float *f;
    } location;
    default_type_t type;
    int bound;
} default_t;

typedef struct
{
    default_t *defaults;
    int numdefaults;
} default_collection_t;

static default_t defaults[] =
{
    { "mouse_sensitivity", {0}, DEFAULT_INT, 0 },
    { "mouse_acceleration", {0}, DEFAULT_FLOAT, 0 },
    { "mouse_threshold", {0}, DEFAULT_INT, 0 },
    { "entry_03", {0}, DEFAULT_INT, 0 },
    { "entry_04", {0}, DEFAULT_INT, 0 },
    { "entry_05", {0}, DEFAULT_INT, 0 },
    { "entry_06", {0}, DEFAULT_INT, 0 },
    { "entry_07", {0}, DEFAULT_INT, 0 },
    { "entry_08", {0}, DEFAULT_INT, 0 },
    { "entry_09", {0}, DEFAULT_INT, 0 },
    { "entry_10", {0}, DEFAULT_INT, 0 },
    { "entry_11", {0}, DEFAULT_INT, 0 },
    { "entry_12", {0}, DEFAULT_INT, 0 },
    { "entry_13", {0}, DEFAULT_INT, 0 },
    { "entry_14", {0}, DEFAULT_INT, 0 },
    { "entry_15", {0}, DEFAULT_INT, 0 },
    { "entry_16", {0}, DEFAULT_INT, 0 },
    { "entry_17", {0}, DEFAULT_INT, 0 },
    { "entry_18", {0}, DEFAULT_INT, 0 },
    { "entry_19", {0}, DEFAULT_INT, 0 },
    { "entry_20", {0}, DEFAULT_INT, 0 },
    { "entry_21", {0}, DEFAULT_INT, 0 },
    { "entry_22", {0}, DEFAULT_INT, 0 },
    { "entry_23", {0}, DEFAULT_INT, 0 },
    { "entry_24", {0}, DEFAULT_INT, 0 },
};

static default_collection_t collection =
{
    defaults,
    arrlen(defaults),
};

int main(void)
{
    return collection.numdefaults != 25;
}
