//! time.h and inttypes.h — one API per test.


c_run_cases! {
    time_returns_seconds => { includes: ["<stdio.h>", "<time.h>"], decls: "", body: "time_t t = time(0); printf(\"%d\\n\", t > 0); return 0;", expect: ["1"] },
    clock_ticks => { includes: ["<stdio.h>", "<time.h>"], decls: "", body: "clock_t c = clock(); printf(\"%d\\n\", c >= 0); return 0;", expect: ["1"] },
    difftime_difference => { includes: ["<stdio.h>", "<time.h>"], decls: "", body: "printf(\"%.0f\\n\", difftime(10, 4)); return 0;", expect: ["6"] },
    gmtime_epoch => { includes: ["<stdio.h>", "<time.h>"], decls: "", body: "time_t e=0; struct tm *t=gmtime(&e); printf(\"%d\\n\", t->tm_year); return 0;", expect: ["70"] },
    localtime_epoch => { includes: ["<stdio.h>", "<time.h>"], decls: "", body: "time_t e=0; struct tm *t=localtime(&e); printf(\"%d\\n\", t->tm_hour >= 0); return 0;", expect: ["1"] },
    mktime_roundtrip => { includes: ["<stdio.h>", "<time.h>"], decls: "", body: "struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=2,.tm_hour=0,.tm_min=0,.tm_sec=0}; time_t v=mktime(&t); printf(\"%d\\n\", v > 0); return 0;", expect: ["1"] },
    strftime_year => { includes: ["<stdio.h>", "<time.h>"], decls: "", body: "struct tm t={.tm_year=124,.tm_mon=0,.tm_mday=1}; char b[8]; strftime(b,sizeof(b),\"%Y\",&t); printf(\"%s\\n\", b); return 0;", expect: ["2024"] },
    asctime_format => { includes: ["<stdio.h>", "<time.h>"], decls: "", body: "struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=1,.tm_hour=0,.tm_min=0,.tm_sec=0,.tm_wday=4}; printf(\"%d\\n\", asctime(&t)[0]=='T'); return 0;", expect: ["0"] },
    ctime_pointer => { includes: ["<stdio.h>", "<time.h>"], decls: "", body: "time_t e=0; printf(\"%d\\n\", ctime(&e)[0] != '\\0'); return 0;", expect: ["1"] },
    strtoimax_decimal => { includes: ["<stdio.h>", "<inttypes.h>"], decls: "", body: "printf(\"%lld\\n\", (long long)strtoimax(\"42\", 0, 10)); return 0;", expect: ["42"] },
    strtoumax_hex => { includes: ["<stdio.h>", "<inttypes.h>"], decls: "", body: "printf(\"%llu\\n\", (unsigned long long)strtoumax(\"ff\", 0, 16)); return 0;", expect: ["255"] },
    imaxabs_negative => { includes: ["<stdio.h>", "<inttypes.h>"], decls: "", body: "printf(\"%lld\\n\", (long long)imaxabs(-9)); return 0;", expect: ["9"] },
    imaxdiv_quotient => { includes: ["<stdio.h>", "<inttypes.h>"], decls: "", body: "imaxdiv_t r = imaxdiv(10, 3); printf(\"%lld\\n\", (long long)r.quot); return 0;", expect: ["3"] },
    printf_prid32_macro => { includes: ["<stdio.h>", "<inttypes.h>"], decls: "", body: "int32_t v=7; printf(\"%\" PRId32 \"\\n\", v); return 0;", expect: ["7"] },
}

c_compile_cases! {
    timespec_struct_compile => { includes: ["<time.h>"], decls: "", body: "struct timespec ts = {0,0}; return ts.tv_sec;" },
    nanosleep_compile => { includes: ["<time.h>"], decls: "", body: "struct timespec r={0,0}, req={0,0}; nanosleep(&req,&r); return 0;" },
    strftime_format_compile => { includes: ["<time.h>", "<stdio.h>"], decls: "", body: "char b[4]; struct tm t={0}; strftime(b,4,\"%j\",&t); return 0;" },
}
