#include <stdio.h>
#include <stdlib.h>

typedef struct Obj Obj;

typedef struct {
    int (*Read)(Obj *obj, int x);
} Klass;

struct Obj {
    Klass *klass;
};

typedef struct {
    Obj obj;
    int payload;
} Wrapped;

static int real_read(Obj *obj, int x)
{
    (void)obj;
    return x + 7;
}

Klass global_klass = {
    real_read,
};

static Obj *make_obj(void)
{
    Wrapped *w = malloc(sizeof(Wrapped));
    w->obj.klass = &global_klass;
    w->payload = 3;
    return &w->obj;
}

int main(void)
{
    Obj *obj = make_obj();
    printf("%d\n", obj->klass->Read(obj, 5));
    return 0;
}
