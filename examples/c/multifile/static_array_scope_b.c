typedef struct {
    int value;
    int spare;
} local_record_t;

static local_record_t records[3];

int second_record_value(void) {
    records[2].value = 17;
    return records[2].value;
}
