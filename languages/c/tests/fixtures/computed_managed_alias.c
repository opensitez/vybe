struct node { int value; };
static struct node original = { 7 };
static int calls;

static struct node *managed_result(void)
{
    ++calls;
    return &original;
}

static struct node record_result(void)
{
    ++calls;
    return original;
}

int main(void)
{
    struct node *cast_alias = (struct node *)(void *)managed_result();
    struct node *sequence_alias = (calls++, &original);
    struct node *call_alias = managed_result();
    struct node copy = record_result();
    if (calls != 4) return 1;
    if (cast_alias != &original || sequence_alias != cast_alias || call_alias != cast_alias) return 2;
    cast_alias->value = 11;
    if (sequence_alias->value != 11 || call_alias->value != 11 || original.value != 11) return 3;
    if (copy.value != 7) return 4;
    copy.value = 19;
    if (original.value != 11) return 5;
    return 0;
}
