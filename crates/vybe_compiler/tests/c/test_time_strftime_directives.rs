//! strftime format directives — one conversion specifier per test with fixed struct tm.


c_run_cases! {
    strftime_percent_y_four_digit_year => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15}; char b[8]; strftime(b,sizeof(b),\"%Y\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["2024"]
    },
    strftime_percent_y_two_digit_year => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=99,.tm_mon=0,.tm_mday=1}; char b[4]; strftime(b,sizeof(b),\"%y\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["99"]
    },
    strftime_percent_m_zero_padded_month => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=0,.tm_mday=1}; char b[4]; strftime(b,sizeof(b),\"%m\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["01"]
    },
    strftime_percent_d_zero_padded_day => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=7}; char b[4]; strftime(b,sizeof(b),\"%d\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["07"]
    },
    strftime_percent_h_twenty_four_hour => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_hour=9}; char b[4]; strftime(b,sizeof(b),\"%H\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["09"]
    },
    strftime_percent_m_zero_padded_minute => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_min=5}; char b[4]; strftime(b,sizeof(b),\"%M\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["05"]
    },
    strftime_percent_s_zero_padded_second => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_sec=3}; char b[4]; strftime(b,sizeof(b),\"%S\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["03"]
    },
    strftime_percent_a_full_weekday_name => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_wday=6}; char b[16]; strftime(b,sizeof(b),\"%A\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["Saturday"]
    },
    strftime_percent_b_full_month_name => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=1}; char b[16]; strftime(b,sizeof(b),\"%B\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["June"]
    },
    strftime_percent_a_abbrev_weekday => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_wday=6}; char b[8]; strftime(b,sizeof(b),\"%a\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["Sat"]
    },
    strftime_percent_b_abbrev_month => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=1}; char b[8]; strftime(b,sizeof(b),\"%b\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["Jun"]
    },
    strftime_percent_p_am_pm_marker => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_hour=14}; char b[4]; strftime(b,sizeof(b),\"%p\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["PM"]
    },
    strftime_percent_i_twelve_hour_clock => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_hour=14}; char b[4]; strftime(b,sizeof(b),\"%I\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["02"]
    },
    strftime_percent_j_day_of_year => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_yday=166}; char b[8]; strftime(b,sizeof(b),\"%j\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["167"]
    },
    strftime_percent_w_weekday_sunday_zero => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_wday=6}; char b[4]; strftime(b,sizeof(b),\"%w\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["6"]
    },
    strftime_percent_u_iso_weekday => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_wday=6}; char b[4]; strftime(b,sizeof(b),\"%u\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["6"]
    },
    strftime_percent_c_century => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=0,.tm_mday=1}; char b[4]; strftime(b,sizeof(b),\"%C\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["20"]
    },
    strftime_percent_f_iso_date => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15}; char b[16]; strftime(b,sizeof(b),\"%F\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["2024-06-15"]
    },
    strftime_percent_d_us_date => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15}; char b[16]; strftime(b,sizeof(b),\"%D\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["06/15/24"]
    },
    strftime_percent_r_hour_minute => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_hour=14,.tm_min=30}; char b[8]; strftime(b,sizeof(b),\"%R\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["14:30"]
    },
    strftime_percent_t_hour_minute_second => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_hour=14,.tm_min=30,.tm_sec=45}; char b[16]; strftime(b,sizeof(b),\"%T\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["14:30:45"]
    },
    strftime_percent_e_space_padded_day => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=7}; char b[4]; strftime(b,sizeof(b),\"%e\",&t); printf(\"%s\\n\", b); return 0;",
        expect: [" 7"]
    },
    strftime_percent_l_space_padded_twelve_hour => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_hour=9}; char b[4]; strftime(b,sizeof(b),\"%l\",&t); printf(\"%s\\n\", b); return 0;",
        expect: [" 9"]
    },
    strftime_percent_k_space_padded_hour => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_hour=9}; char b[4]; strftime(b,sizeof(b),\"%k\",&t); printf(\"%s\\n\", b); return 0;",
        expect: [" 9"]
    },
    strftime_percent_h_same_as_abbrev_month => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=11,.tm_mday=1}; char b[8]; strftime(b,sizeof(b),\"%h\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["Dec"]
    },
    strftime_double_percent_literal => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={0}; char b[8]; strftime(b,sizeof(b),\"%%\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["%"]
    },
    strftime_midnight_am => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=0,.tm_mday=1,.tm_hour=0}; char b[4]; strftime(b,sizeof(b),\"%p\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["AM"]
    },
    strftime_noon_pm => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=0,.tm_mday=1,.tm_hour=12}; char b[4]; strftime(b,sizeof(b),\"%p\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["PM"]
    },
    strftime_leap_day_february => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=1,.tm_mday=29}; char b[16]; strftime(b,sizeof(b),\"%Y-%m-%d\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["2024-02-29"]
    },
    strftime_year_boundary_december => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=123,.tm_mon=11,.tm_mday=31}; char b[16]; strftime(b,sizeof(b),\"%Y-%m-%d\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["2023-12-31"]
    },
    strftime_january_first_yday => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=0,.tm_mday=1,.tm_yday=0}; char b[8]; strftime(b,sizeof(b),\"%j\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["001"]
    },
    strftime_december_thirty_first_yday => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=123,.tm_mon=11,.tm_mday=31,.tm_yday=364}; char b[8]; strftime(b,sizeof(b),\"%j\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["365"]
    },
    strftime_sunday_wday_zero => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=0,.tm_mday=7,.tm_wday=0}; char b[16]; strftime(b,sizeof(b),\"%A\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["Sunday"]
    },
    strftime_monday_wday_one => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=0,.tm_mday=8,.tm_wday=1}; char b[16]; strftime(b,sizeof(b),\"%A\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["Monday"]
    },
    strftime_march_month_name => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=2,.tm_mday=1}; char b[16]; strftime(b,sizeof(b),\"%B\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["March"]
    },
    strftime_october_abbrev => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=9,.tm_mday=1}; char b[8]; strftime(b,sizeof(b),\"%b\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["Oct"]
    },
    strftime_late_night_hour => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_hour=23,.tm_min=59,.tm_sec=59}; char b[16]; strftime(b,sizeof(b),\"%H:%M:%S\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["23:59:59"]
    },
    strftime_early_morning_twelve_hour => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_hour=1}; char b[8]; strftime(b,sizeof(b),\"%I %p\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["01 AM"]
    },
    strftime_century_nineteen => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=99,.tm_mon=6,.tm_mday=4}; char b[4]; strftime(b,sizeof(b),\"%C\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["19"]
    },
    strftime_combined_ymd_hms => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=1,.tm_hour=0,.tm_min=0,.tm_sec=0}; char b[32]; strftime(b,sizeof(b),\"%Y%m%d%H%M%S\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["19700101000000"]
    },
    strftime_buffer_truncation_returns_length => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15}; char b[4]; size_t n=strftime(b,sizeof(b),\"%Y-%m-%d\",&t); printf(\"%zu\\n\", n); return 0;",
        expect: ["10"]
    },
    strftime_empty_format_writes_nul => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15}; char b[4]={'X','X','X','X'}; strftime(b,sizeof(b),\"\",&t); printf(\"%d\\n\", b[0]==0); return 0;",
        expect: ["1"]
    },
    strftime_percent_v_iso_week => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_wday=6}; char b[4]; strftime(b,sizeof(b),\"%V\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["24"]
    },
    strftime_percent_g_iso_week_year => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_wday=6}; char b[8]; strftime(b,sizeof(b),\"%G\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["2024"]
    },
    strftime_percent_g_iso_week_year_two_digit => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_wday=6}; char b[4]; strftime(b,sizeof(b),\"%g\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["24"]
    },
    strftime_percent_u_sunday_week_number => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_wday=6,.tm_yday=166}; char b[4]; strftime(b,sizeof(b),\"%U\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["23"]
    },
    strftime_percent_w_monday_week_number => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_wday=6,.tm_yday=166}; char b[4]; strftime(b,sizeof(b),\"%W\",&t); printf(\"%s\\n\", b); return 0;",
        expect: ["24"]
    },
}

c_compile_cases! {
    strftime_percent_n_newline_compile => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={0}; char b[4]; strftime(b,sizeof(b),\"%n\",&t); return 0;"
    },
    strftime_percent_t_tab_compile => {
        includes: ["<stdio.h>", "<time.h>"],
        decls: "",
        body: "struct tm t={0}; char b[4]; strftime(b,sizeof(b),\"%t\",&t); return 0;"
    },
}
