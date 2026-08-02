// vybe-test: js/symbol_protocols/symbol_as_constant_enum
// origin: languages/js/tests/js/test_symbol_protocols.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

const Direction = {
    UP: Symbol("UP"),
    DOWN: Symbol("DOWN"),
    LEFT: Symbol("LEFT"),
    RIGHT: Symbol("RIGHT"),
};
function move(dir) {
    switch(dir) {
        case Direction.UP: return "going up";
        case Direction.DOWN: return "going down";
        default: return "other";
    }
}
__check(__line(move(Direction.UP)), "going up");
__check(__line(move(Direction.DOWN)), "going down");
__check(__line(move(Direction.LEFT)), "other");
__check(__line(Direction.UP !== Direction.DOWN), "true");
