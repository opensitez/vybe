#include <assert.h>

typedef struct {
    int value;
    int spare;
} local_record_t;

static local_record_t records[32];

int second_record_value(void);

int first_record_value(void) {
    records[3].value = 41;
    return records[3].value;
}

int main(void) {
    assert(first_record_value() == 41);
    assert(second_record_value() == 17);
    assert(records[3].value == 41);
    return 0;
}
