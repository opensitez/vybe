//! `os` library: `os.clock`, `os.time`, `os.date`, `os.difftime` (Lua 5.x §6.9)

lua_print! {
os_clock_type => { "print(type(os.clock()))\n", "number" },
os_clock_nonneg => { "print(os.clock() >= 0)\n", "true" },
os_time_type => { "print(type(os.time()))\n", "number" },
os_time_pos => { "print(os.time() > 0)\n", "true" },
os_time_epoch => {
    "local t = os.time({year=2000, month=1, day=1, hour=0, min=0, sec=0})\nprint(t > 0)\n",
    "true"
},
os_difftime_check => {
    "local t1 = os.time({year=2000, month=1, day=1, hour=0, min=0, sec=0})\nlocal t2 = os.time({year=2000, month=1, day=1, hour=0, min=0, sec=5})\nprint(os.difftime(t2, t1))\n",
    "5.0"
},
os_date_default => {
    "print(type(os.date()))\n",
    "string"
},
os_date_year => {
    "local epoch = os.time({year=2024, month=6, day=15, hour=12, min=0, sec=0})\nprint(os.date(\"%Y\", epoch))\n",
    "2024"
},
os_date_month => {
    "local epoch = os.time({year=2024, month=6, day=15, hour=12, min=0, sec=0})\nprint(os.date(\"%m\", epoch))\n",
    "06"
},
os_date_day => {
    "local epoch = os.time({year=2024, month=6, day=15, hour=12, min=0, sec=0})\nprint(os.date(\"%d\", epoch))\n",
    "15"
},
os_date_table => {
    "local epoch = os.time({year=2024, month=3, day=10, hour=0, min=0, sec=0})\nlocal d = os.date(\"*t\", epoch)\nprint(d.year .. \",\" .. d.month .. \",\" .. d.day)\n",
    "2024,3,10"
},
os_time_ordering => {
    "local earlier = os.time({year=2000, month=1, day=1, hour=0, min=0, sec=0})\nlocal later = os.time({year=2001, month=1, day=1, hour=0, min=0, sec=0})\nprint(later > earlier)\n",
    "true"
} }
