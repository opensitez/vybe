//! `NumberFormatter`, `IntlDateFormatter`, and locale helpers when ext-intl is available.

crate::php_cases! {
    intl_numberformatter_decimal_when_available => {
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skip'; } else {
    $fmt = new NumberFormatter('en_US', NumberFormatter::DECIMAL);
    echo $fmt->format(1234.5);
}
"#,
        ["1,234.5"]
    };

    intl_numberformatter_currency_usd => {
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skip'; } else {
    $fmt = new NumberFormatter('en_US', NumberFormatter::CURRENCY);
    echo str_contains($fmt->formatCurrency(9.5, 'USD'), '9.50') ? 'usd' : 'no';
}
"#,
        ["usd"]
    };

    intl_numberformatter_percent => {
        r#"<?php
if (!class_exists('NumberFormatter')) { echo 'skip'; } else {
    $fmt = new NumberFormatter('en_US', NumberFormatter::PERCENT);
    echo str_contains($fmt->format(0.25), '25') ? 'pct' : 'no';
}
"#,
        ["pct"]
    };

    intl_dateformatter_short_date => {
        r#"<?php
if (!class_exists('IntlDateFormatter')) { echo 'skip'; } else {
    date_default_timezone_set('UTC');
    $fmt = new IntlDateFormatter('en_US', IntlDateFormatter::SHORT, IntlDateFormatter::NONE, 'UTC');
    echo strlen($fmt->format(1704067200)) > 0 ? 'dated' : 'empty';
}
"#,
        ["dated"]
    };

    locale_get_default_returns_string => {
        r#"<?php
echo is_string(locale_get_default()) ? 'loc' : 'no';
"#,
        ["loc"]
    };

    intl_get_error_code_zero_when_no_error => {
        r#"<?php
if (!function_exists('intl_get_error_code')) { echo 'skip'; } else {
    echo intl_get_error_code() === 0 ? 'ok' : 'err';
}
"#,
        ["ok"]
    };

    resourcebundle_create_when_available => {
        r#"<?php
if (!class_exists('ResourceBundle')) { echo 'skip'; } else {
    $rb = ResourceBundle::create('en', null, false);
    echo $rb !== null ? 'bundle' : 'null';
}
"#,
        ["bundle"]
    };

    collator_compare_orders_strings => {
        r#"<?php
if (!class_exists('Collator')) { echo 'skip'; } else {
    $c = new Collator('en_US');
    echo $c->compare('a', 'b') < 0 ? 'asc' : 'desc';
}
"#,
        ["asc"]
    };

    normalizer_normalize_nfc => {
        r#"<?php
if (!class_exists('Normalizer')) { echo 'skip'; } else {
    $n = Normalizer::normalize("e\u{0301}", Normalizer::FORM_C);
    echo is_string($n) ? 'norm' : 'fail';
}
"#,
        ["norm"]
    };

    transliterator_transliterate_when_available => {
        r#"<?php
if (!class_exists('Transliterator')) { echo 'skip'; } else {
    $tr = Transliterator::create('Any-Latin; Latin-ASCII');
    echo $tr ? 'tr' : 'no';
}
"#,
        ["tr"]
    };
}
