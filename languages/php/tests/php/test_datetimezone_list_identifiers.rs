
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

    datetimezone_list_all_contains_utc => {
        r#"<?php
$zones = DateTimeZone::listIdentifiers();
echo in_array('UTC', $zones) ? "utc" : "no_utc";
"#,
        ["utc"]
    };

    datetimezone_list_identifiers_invalid_country => {
        r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::PER_COUNTRY, 'ZZ');
echo count($zones) === 0 ? "none" : "some";
"#,
        ["none"]
    };

    datetimezone_list_identifiers_country_hint_counts => {
        r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::PER_COUNTRY, 'US');
echo is_array($zones) ? (count($zones) > 0 ? "has_us" : "no_us") : "bad";
"#,
        ["has_us"]
    };

    datetimezone_list_identifiers_america_region => {
        r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::AMERICA);
echo count($zones) > 1 ? "many" : "few";
"#,
        ["many"]
    };

    datetimezone_list_identifiers_africa_region => {
        r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::AFRICA);
echo is_array($zones) ? (count($zones) > 0 ? "yes" : "no") : "bad";
echo ':' . (in_array('Africa/Cairo', $zones) ? 'cairo' : 'nocairo');
"#,
        ["yes:cairo"]
    };

    datetimezone_list_identifiers_europe_region => {
        r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::EUROPE);
echo count($zones) > 10 ? "many" : "few";
echo ':';
echo in_array('Europe/Paris', $zones) ? 'paris' : 'noparis';
"#,
        ["many:paris"]
    };

    datetimezone_list_identifiers_asiapacific => {
        r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::ASIA);
echo count($zones) > 0 ? "found" : "none";
echo ':';
echo in_array('Asia/Tokyo', $zones) ? 'tokyo' : 'notokyo';
"#,
        ["found:tokyo"]
    };

    datetimezone_list_identifiers_antarctica => {
        r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::ANTARCTICA);
echo is_array($zones) ? 'arr' : 'bad';
echo ':' . (count($zones) > 0 ? 'some' : 'none');
"#,
        ["arr:some"]
    };
}
