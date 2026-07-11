//! Epoch and calendar arithmetic — difftime, mktime normalization, time_t ordering.

c_run_cases! {
    difftime_later_minus_earlier_positive => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", difftime(5000, 1000)); return 0;",
        expect: ["4000"]
    },
    difftime_earlier_minus_later_negative => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", difftime(100, 500)); return 0;",
        expect: ["-400"]
    },
    difftime_equal_times_zero => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", difftime(777, 777)); return 0;",
        expect: ["0"]
    },
    difftime_one_second_apart => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", difftime(1, 0)); return 0;",
        expect: ["1"]
    },
    difftime_epoch_to_thousand => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", difftime(1000, 0)); return 0;",
        expect: ["1000"]
    },
    difftime_large_span => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", difftime(1000000, 1)); return 0;",
        expect: ["999999"]
    },
    difftime_negative_epoch_span => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", difftime(-100, -1100)); return 0;",
        expect: ["1000"]
    },
    difftime_crossing_zero => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", difftime(50, -50)); return 0;",
        expect: ["100"]
    },
    time_t_less_than_comparison => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t a=100,b=200; printf(\"%d\\n\", a<b); return 0;",
        expect: ["1"]
    },
    time_t_greater_than_comparison => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t a=300,b=200; printf(\"%d\\n\", a>b); return 0;",
        expect: ["1"]
    },
    time_t_equality_same_instant => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t a=42,b=42; printf(\"%d\\n\", a==b); return 0;",
        expect: ["1"]
    },
    time_t_inequality_different_instant => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t a=42,b=43; printf(\"%d\\n\", a!=b); return 0;",
        expect: ["1"]
    },
    time_t_less_equal_when_equal => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t a=99,b=99; printf(\"%d\\n\", a<=b); return 0;",
        expect: ["1"]
    },
    time_t_greater_equal_when_greater => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t a=101,b=100; printf(\"%d\\n\", a>=b); return 0;",
        expect: ["1"]
    },
    mktime_epoch_jan_first_1970 => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=1,.tm_hour=0,.tm_min=0,.tm_sec=0}; time_t v=mktime(&t); printf(\"%lld\\n\", (long long)v); return 0;",
        expect: ["0"]
    },
    mktime_one_day_after_epoch => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=2,.tm_hour=0,.tm_min=0,.tm_sec=0}; time_t v=mktime(&t); printf(\"%lld\\n\", (long long)v); return 0;",
        expect: ["86400"]
    },
    mktime_seconds_overflow_into_minute => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=1,.tm_hour=0,.tm_min=0,.tm_sec=75}; mktime(&t); printf(\"%d %d\\n\", t.tm_min, t.tm_sec); return 0;",
        expect: ["1 15"]
    },
    mktime_minutes_overflow_into_hour => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=1,.tm_hour=0,.tm_min=90,.tm_sec=0}; mktime(&t); printf(\"%d %d\\n\", t.tm_hour, t.tm_min); return 0;",
        expect: ["1 30"]
    },
    mktime_hours_overflow_into_day => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=1,.tm_hour=25,.tm_min=0,.tm_sec=0}; mktime(&t); printf(\"%d %d\\n\", t.tm_mday, t.tm_hour); return 0;",
        expect: ["2 1"]
    },
    mktime_day_overflow_into_month => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=32,.tm_hour=0,.tm_min=0,.tm_sec=0}; mktime(&t); printf(\"%d %d\\n\", t.tm_mon+1, t.tm_mday); return 0;",
        expect: ["2 1"]
    },
    mktime_month_overflow_into_year => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=70,.tm_mon=12,.tm_mday=1,.tm_hour=0,.tm_min=0,.tm_sec=0}; mktime(&t); printf(\"%d %d\\n\", t.tm_year+1900, t.tm_mon+1); return 0;",
        expect: ["1971 1"]
    },
    mktime_negative_month_rolls_back_year => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=71,.tm_mon=-1,.tm_mday=15,.tm_hour=0,.tm_min=0,.tm_sec=0}; mktime(&t); printf(\"%d %d\\n\", t.tm_year+1900, t.tm_mon+1); return 0;",
        expect: ["1970 12"]
    },
    mktime_leap_day_february_twenty_ninth => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=1,.tm_mday=29,.tm_hour=12,.tm_min=0,.tm_sec=0}; mktime(&t); printf(\"%d %d\\n\", t.tm_mon+1, t.tm_mday); return 0;",
        expect: ["2 29"]
    },
    mktime_march_first_after_leap_feb => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=2,.tm_mday=1,.tm_hour=0,.tm_min=0,.tm_sec=0}; time_t v=mktime(&t); printf(\"%lld\\n\", (long long)difftime(v, 0) > 1700000000); return 0;",
        expect: ["1"]
    },
    mktime_december_thirty_first => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=123,.tm_mon=11,.tm_mday=31,.tm_hour=23,.tm_min=59,.tm_sec=59}; mktime(&t); printf(\"%d\\n\", t.tm_year+1900); return 0;",
        expect: ["2023"]
    },
    mktime_yday_updated_after_normalize => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_hour=0,.tm_min=0,.tm_sec=0}; mktime(&t); printf(\"%d\\n\", t.tm_yday); return 0;",
        expect: ["166"]
    },
    mktime_wday_computed_for_known_date => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_hour=0,.tm_min=0,.tm_sec=0}; mktime(&t); printf(\"%d\\n\", t.tm_wday); return 0;",
        expect: ["6"]
    },
    gmtime_epoch_fields => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t e=0; struct tm *p=gmtime(&e); printf(\"%d %d %d\\n\", p->tm_year+1900, p->tm_mon+1, p->tm_mday); return 0;",
        expect: ["1970 1 1"]
    },
    gmtime_one_hour_offset => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t e=3600; struct tm *p=gmtime(&e); printf(\"%d\\n\", p->tm_hour); return 0;",
        expect: ["1"]
    },
    gmtime_day_rollover_at_midnight_utc => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t e=86400; struct tm *p=gmtime(&e); printf(\"%d %d\\n\", p->tm_mday, p->tm_hour); return 0;",
        expect: ["2 0"]
    },
    difftime_matches_time_t_subtraction_magnitude => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t a=2500,b=1500; printf(\"%.0f\\n\", difftime(a,b)); return 0;",
        expect: ["1000"]
    },
    mktime_roundtrip_preserves_calendar_day => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=120,.tm_mon=6,.tm_mday=4,.tm_hour=15,.tm_min=30,.tm_sec=0}; time_t v=mktime(&t); struct tm *p=gmtime(&v); printf(\"%d %d\\n\", p->tm_mon+1, p->tm_mday); return 0;",
        expect: ["7 4"]
    },
    mktime_negative_seconds_normalize => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=1,.tm_hour=1,.tm_min=0,.tm_sec=-30}; mktime(&t); printf(\"%d %d\\n\", t.tm_hour, t.tm_sec); return 0;",
        expect: ["0 30"]
    },
    mktime_combined_overflow_chain => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=1,.tm_hour=23,.tm_min=59,.tm_sec=120}; mktime(&t); printf(\"%d %d %d\\n\", t.tm_mday, t.tm_hour, t.tm_min); return 0;",
        expect: ["2 0 1"]
    },
    time_t_zero_is_epoch => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t t=0; printf(\"%d\\n\", t==0); return 0;",
        expect: ["1"]
    },
    difftime_from_epoch_to_known_offset => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", difftime(172800, 0)); return 0;",
        expect: ["172800"]
    },
    mktime_april_thirtieth_non_leap => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=123,.tm_mon=3,.tm_mday=30,.tm_hour=0,.tm_min=0,.tm_sec=0}; mktime(&t); printf(\"%d %d\\n\", t.tm_mon+1, t.tm_mday); return 0;",
        expect: ["4 30"]
    },
    mktime_may_first_from_april_overflow => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=123,.tm_mon=3,.tm_mday=31,.tm_hour=0,.tm_min=0,.tm_sec=0}; mktime(&t); printf(\"%d %d\\n\", t.tm_mon+1, t.tm_mday); return 0;",
        expect: ["5 1"]
    },
    mktime_hour_twenty_three_fifty_nine => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=0,.tm_mday=1,.tm_hour=23,.tm_min=59,.tm_sec=0}; mktime(&t); printf(\"%d\\n\", t.tm_hour*100+t.tm_min); return 0;",
        expect: ["2359"]
    },
    gmtime_year_two_thousand => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t e=946684800; struct tm *p=gmtime(&e); printf(\"%d\\n\", p->tm_year+1900); return 0;",
        expect: ["2000"]
    },
    difftime_subsecond_truncates_to_whole_seconds => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", difftime(10, 3)); return 0;",
        expect: ["7"]
    },
    mktime_isdst_preserved_field => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=6,.tm_mday=1,.tm_hour=12,.tm_isdst=1}; mktime(&t); printf(\"%d\\n\", t.tm_isdst); return 0;",
        expect: ["1"]
    },
    time_t_compare_chain_ordering => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t a=10,b=20,c=30; printf(\"%d\\n\", a<b && b<c); return 0;",
        expect: ["1"]
    },
    mktime_january_thirty_first => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=0,.tm_mday=31,.tm_hour=0,.tm_min=0,.tm_sec=0}; mktime(&t); printf(\"%d %d\\n\", t.tm_mon+1, t.tm_mday); return 0;",
        expect: ["1 31"]
    },
    mktime_february_first_from_jan_overflow => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=0,.tm_mday=32,.tm_hour=0,.tm_min=0,.tm_sec=0}; mktime(&t); printf(\"%d %d\\n\", t.tm_mon+1, t.tm_mday); return 0;",
        expect: ["2 1"]
    },
    difftime_symmetry_negated => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "double a=difftime(800,300); double b=difftime(300,800); printf(\"%.0f\\n\", a+b); return 0;",
        expect: ["0"]
    },
    gmtime_sec_field_at_offset => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "time_t e=45; struct tm *p=gmtime(&e); printf(\"%d\\n\", p->tm_sec); return 0;",
        expect: ["45"]
    },
}

c_compile_cases! {
    time_t_type_compiles => { includes: ["<time.h>"], decls: "", body: "time_t t=0; return (int)t;" },
    difftime_signature_compiles => { includes: ["<time.h>"], decls: "", body: "return (int)difftime(1,0);" },
}
