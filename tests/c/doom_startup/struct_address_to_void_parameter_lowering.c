// vybe-test: c/doom_startup/struct_address_to_void_parameter_lowering
typedef unsigned long size_t;

typedef struct
{
    char identification[4];
    int numlumps;
    int infotableofs;
} wadinfo_t;

static size_t accept_void(void *buffer)
{
    return 0;
}

int main(void)
{
    wadinfo_t header;
    return (int) accept_void(&header);
}
