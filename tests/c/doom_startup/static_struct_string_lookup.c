int strcmp(const char *a, const char *b);

typedef struct
{
    const char *name;
    int value;
} entry_t;

typedef struct
{
    entry_t *entries;
    int count;
} collection_t;

static entry_t defaults[] =
{
    { "mouse_sensitivity", 1 },
    { "mouse_acceleration", 2 },
    { "mouse_threshold", 3 },
};

static collection_t collection =
{
    defaults,
    3,
};

static entry_t *find(collection_t *collection, const char *name)
{
    int i;

    for (i = 0; i < collection->count; ++i)
    {
        if (!strcmp(name, collection->entries[i].name))
        {
            return &collection->entries[i];
        }
    }

    return 0;
}

int main(void)
{
    entry_t *entry;

    entry = find(&collection, "mouse_acceleration");
    if (entry == 0)
    {
        return 1;
    }

    return entry->value != 2;
}
