use super::helpers::run_prints;

crate::php_cases! {
    datetimezone_list_identifiers => {
        r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::EUROPE);
echo in_array('Europe/London', $zones) ? "found" : "missing";
"#,
        ["found"]
    };

    datetimezone_list_identifiers_per_country => {
        r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::PER_COUNTRY, 'GB');
echo implode(',', $zones);
"#,
        ["Europe/London"]
    };
}
