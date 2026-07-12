//! math.random and math.randomseed boundary logic (Lua 5.x §6.7)

lua_print! {
    random_no_args => {
        "math.randomseed(12345)\nlocal r = math.random()\nprint(r >= 0 and r < 1)\n",
        "true"
    },
    random_one_arg => {
        "math.randomseed(12345)\nlocal r = math.random(10)\nprint(r >= 1 and r <= 10 and math.type(r) == \"integer\")\n",
        "true"
    },
    random_two_args => {
        "math.randomseed(12345)\nlocal r = math.random(5, 15)\nprint(r >= 5 and r <= 15 and math.type(r) == \"integer\")\n",
        "true"
    },
    random_reproducible => {
        "math.randomseed(42)\nlocal r1 = math.random(1000)\nmath.randomseed(42)\nlocal r2 = math.random(1000)\nprint(r1 == r2)\n",
        "true"
    },
    random_equal_bounds => {
        "print(math.random(7, 7))\n",
        "7"
    },
}
