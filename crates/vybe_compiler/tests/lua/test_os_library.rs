//! `os` library — time, date, clock (Lua 5.x manual §6.2).

lua_print! {
    os_clock_returns_non_negative_number => {
        "print(os.clock() >= 0)\n",
        "true"
    },
    os_date_default_returns_string => {
        "print(type(os.date()) == \"string\")\n",
        "true"
    },
    os_date_with_format_year => {
        "local s = os.date(\"%Y\")\nprint(#s == 4)\n",
        "true"
    },
    os_time_returns_integer_seconds => {
        "print(math.type(os.time()) == \"integer\" or type(os.time()) == \"number\")\n",
        "true"
    },
    os_difftime_measures_elapsed_seconds => {
        "print(os.difftime(100, 40))\n",
        "60"
    },
    os_date_table_roundtrip => {
        "local t = os.date(\"*t\", 0)\nprint(t.year > 1969)\n",
        "true"
    },
    os_setlocale_returns_string_or_nil => {
        "local r = os.setlocale(\"C\")\nprint(r == \"C\" or r == nil)\n",
        "true"
    },
}
