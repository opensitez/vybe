//! locale.h — one distinct API or constant per test.

use crate::helpers::*;

c_compile_cases! {
    locale_setlocale_c => { includes: ["<locale.h>"], decls: "", body: "setlocale(LC_ALL, \"C\"); return 0;" },
    locale_setlocale_null => { includes: ["<locale.h>"], decls: "", body: "return setlocale(LC_ALL, 0) != 0;" },
    localeconv_decimal => { includes: ["<locale.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->decimal_point[0];" },
    localeconv_thousands => { includes: ["<locale.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->thousands_sep[0] || 1;" },
    localeconv_grouping => { includes: ["<locale.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->grouping != 0;" },
    localeconv_int_curr_symbol => { includes: ["<locale.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->int_curr_symbol != 0;" },
    localeconv_currency_symbol => { includes: ["<locale.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->currency_symbol != 0;" },
    localeconv_mon_decimal => { includes: ["<locale.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->mon_decimal_point != 0;" },
    localeconv_mon_thousands => { includes: ["<locale.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->mon_thousands_sep != 0;" },
    localeconv_mon_grouping => { includes: ["<locale.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->mon_grouping != 0;" },
    localeconv_positive_sign => { includes: ["<locale.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->positive_sign != 0;" },
    localeconv_negative_sign => { includes: ["<locale.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->negative_sign != 0;" },
    localeconv_int_frac_digits => { includes: ["<locale.h>", "<limits.h>"], decls: "", body: "struct lconv *lc = localeconv(); return (int)lc->int_frac_digits >= 0 || lc->int_frac_digits == CHAR_MAX;" },
    localeconv_frac_digits => { includes: ["<locale.h>", "<limits.h>"], decls: "", body: "struct lconv *lc = localeconv(); return (int)lc->frac_digits >= 0 || lc->frac_digits == CHAR_MAX;" },
    localeconv_p_cs_precedes => { includes: ["<locale.h>", "<limits.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->p_cs_precedes == CHAR_MAX || lc->p_cs_precedes >= 0;" },
    localeconv_p_sep_by_space => { includes: ["<locale.h>", "<limits.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->p_sep_by_space == CHAR_MAX || lc->p_sep_by_space >= 0;" },
    localeconv_p_sign_posn => { includes: ["<locale.h>", "<limits.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->p_sign_posn == CHAR_MAX || lc->p_sign_posn >= 0;" },
    localeconv_n_cs_precedes => { includes: ["<locale.h>", "<limits.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->n_cs_precedes == CHAR_MAX || lc->n_cs_precedes >= 0;" },
    localeconv_n_sep_by_space => { includes: ["<locale.h>", "<limits.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->n_sep_by_space == CHAR_MAX || lc->n_sep_by_space >= 0;" },
    localeconv_n_sign_posn => { includes: ["<locale.h>", "<limits.h>"], decls: "", body: "struct lconv *lc = localeconv(); return lc->n_sign_posn == CHAR_MAX || lc->n_sign_posn >= 0;" },
    lc_all_constant => { includes: ["<locale.h>"], decls: "", body: "return LC_ALL != 0;" },
    lc_collate_constant => { includes: ["<locale.h>"], decls: "", body: "return LC_COLLATE != 0;" },
    lc_ctype_constant => { includes: ["<locale.h>"], decls: "", body: "return LC_CTYPE != 0;" },
    lc_monetary_constant => { includes: ["<locale.h>"], decls: "", body: "return LC_MONETARY != 0;" },
    lc_numeric_constant => { includes: ["<locale.h>"], decls: "", body: "return LC_NUMERIC != 0;" },
    lc_time_constant => { includes: ["<locale.h>"], decls: "", body: "return LC_TIME != 0;" },
    lc_messages_if_defined => { includes: ["<locale.h>"], decls: "#ifdef LC_MESSAGES\nint use_lc_messages(void){return LC_MESSAGES;}\n#endif", body: "return 0;" },
}
