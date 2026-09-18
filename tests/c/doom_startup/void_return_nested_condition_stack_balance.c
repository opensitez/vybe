typedef struct {
    int special;
} line_t;

typedef struct {
    int player;
    int type;
} mobj_t;

void cross(line_t *line, mobj_t *thing) {
    if (line->special > 98 && line->special != 104) {
        return;
    }

    if (!thing->player) {
        switch (thing->type) {
            case 1:
            case 2:
                return;
            default:
                break;
        }
    }
}

int main(void) {
    line_t line;
    mobj_t thing;
    line.special = 99;
    thing.player = 0;
    thing.type = 1;
    cross(&line, &thing);
    return 0;
}
