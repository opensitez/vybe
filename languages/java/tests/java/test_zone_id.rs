use crate::helpers::run_main;

#[test]
fn zone_id_of_utc_get_id() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("UTC").getId());"#);
    assert_eq!(out, vec!["UTC"]);
}

#[test]
fn zone_id_of_gmt_get_id() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("GMT").getId());"#);
    assert_eq!(out, vec!["GMT"]);
}

#[test]
fn zone_id_of_z_get_id() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("Z").getId());"#);
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn zone_id_of_europe_paris() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("Europe/Paris").getId());"#);
    assert_eq!(out, vec!["Europe/Paris"]);
}

#[test]
fn zone_id_of_america_new_york() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("America/New_York").getId());"#);
    assert_eq!(out, vec!["America/New_York"]);
}

#[test]
fn zone_id_of_asia_tokyo() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("Asia/Tokyo").getId());"#);
    assert_eq!(out, vec!["Asia/Tokyo"]);
}

#[test]
fn zone_id_of_offset_plus_two() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("+02:00").getId());"#);
    assert_eq!(out, vec!["+02:00"]);
}

#[test]
fn zone_id_of_offset_minus_five() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("-05:00").getId());"#);
    assert_eq!(out, vec!["-05:00"]);
}

#[test]
fn zone_id_of_offset_plus_zero() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("+00:00").getId());"#);
    assert_eq!(out, vec!["+00:00"]);
}

#[test]
fn zone_id_system_default_not_null() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.systemDefault() != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zone_id_equals_same_region() {
    let out = run_main(
        r#"java.time.ZoneId a = java.time.ZoneId.of("UTC"); java.time.ZoneId b = java.time.ZoneId.of("UTC"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zone_id_equals_different_region() {
    let out = run_main(
        r#"java.time.ZoneId a = java.time.ZoneId.of("UTC"); java.time.ZoneId b = java.time.ZoneId.of("GMT"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn zone_id_hash_code_consistent() {
    let out = run_main(
        r#"java.time.ZoneId a = java.time.ZoneId.of("UTC"); java.time.ZoneId b = java.time.ZoneId.of("UTC"); System.out.println(a.hashCode() == b.hashCode());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zone_id_to_string_returns_id() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("Europe/London").toString());"#);
    assert_eq!(out, vec!["Europe/London"]);
}

#[test]
fn zone_id_of_offset_hours_only() {
    let out = run_main(
        r#"System.out.println(java.time.ZoneId.ofOffset("Z", java.time.ZoneOffset.ofHours(3)).getId());"#,
    );
    assert_eq!(out, vec!["+03:00"]);
}

#[test]
fn zone_id_of_offset_negative_hours() {
    let out = run_main(
        r#"System.out.println(java.time.ZoneId.ofOffset("Z", java.time.ZoneOffset.ofHours(-8)).getId());"#,
    );
    assert_eq!(out, vec!["-08:00"]);
}

#[test]
fn zone_id_get_rules_not_null() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("UTC").getRules() != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zone_id_short_ids_contains_utc() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.SHORT_IDS.containsKey("UTC"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zone_id_short_ids_maps_est() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.SHORT_IDS.get("EST"));"#);
    assert_eq!(out, vec!["America/New_York"]);
}

#[test]
fn zone_id_normalized_same_for_utc() {
    let out = run_main(
        r#"java.time.ZoneId z = java.time.ZoneId.of("UTC"); System.out.println(z.normalized().getId());"#,
    );
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn zone_id_from_zone_offset() {
    let out = run_main(
        r#"System.out.println(java.time.ZoneId.from(java.time.ZoneOffset.ofHours(5)).getId());"#,
    );
    assert_eq!(out, vec!["+05:00"]);
}

#[test]
fn zone_id_compare_to_same() {
    let out = run_main(
        r#"java.time.ZoneId a = java.time.ZoneId.of("UTC"); java.time.ZoneId b = java.time.ZoneId.of("UTC"); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn zone_id_compare_to_earlier_offset() {
    let out = run_main(
        r#"java.time.ZoneId a = java.time.ZoneId.of("+01:00"); java.time.ZoneId b = java.time.ZoneId.of("+02:00"); System.out.println(a.compareTo(b) < 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zone_id_at_instant_utc_offset_zero() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2024-06-15T12:00:00Z"); System.out.println(java.time.ZoneId.of("UTC").getRules().getOffset(i).getTotalSeconds());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn zone_id_at_instant_fixed_offset() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2024-06-15T12:00:00Z"); System.out.println(java.time.ZoneId.of("+03:00").getRules().getOffset(i).getTotalSeconds());"#,
    );
    assert_eq!(out, vec!["10800"]);
}

#[test]
fn zone_id_of_australia_sydney() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("Australia/Sydney").getId());"#);
    assert_eq!(out, vec!["Australia/Sydney"]);
}

#[test]
fn zone_id_of_pacific_auckland() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("Pacific/Auckland").getId());"#);
    assert_eq!(out, vec!["Pacific/Auckland"]);
}

#[test]
fn zone_id_of_africa_cairo() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("Africa/Cairo").getId());"#);
    assert_eq!(out, vec!["Africa/Cairo"]);
}

#[test]
fn zone_id_of_offset_half_hour() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("+05:30").getId());"#);
    assert_eq!(out, vec!["+05:30"]);
}

#[test]
fn zone_id_of_offset_quarter_hour() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("+05:45").getId());"#);
    assert_eq!(out, vec!["+05:45"]);
}

#[test]
fn zone_id_from_offset_zero() {
    let out =
        run_main(r#"System.out.println(java.time.ZoneId.from(java.time.ZoneOffset.UTC).getId());"#);
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn zone_id_short_ids_contains_pst() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.SHORT_IDS.containsKey("PST"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zone_id_short_ids_maps_cst() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.SHORT_IDS.get("CST"));"#);
    assert_eq!(out, vec!["America/Chicago"]);
}

#[test]
fn zone_id_of_etc_gmt() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("Etc/GMT").getId());"#);
    assert_eq!(out, vec!["Etc/GMT"]);
}

#[test]
fn zone_id_of_etc_gmt_plus_one() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("Etc/GMT+1").getId());"#);
    assert_eq!(out, vec!["Etc/GMT+1"]);
}

#[test]
fn zone_id_get_display_name_utc() {
    let out = run_main(
        r#"System.out.println(java.time.ZoneId.of("UTC").getDisplayName(java.time.format.TextStyle.SHORT, java.util.Locale.ENGLISH));"#,
    );
    assert_eq!(out, vec!["UTC"]);
}

#[test]
fn zone_id_is_normalised_for_fixed_offset() {
    let out = run_main(
        r#"java.time.ZoneId z = java.time.ZoneId.of("+04:00"); System.out.println(z.equals(z.normalized()));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zone_id_of_minute_precision_offset() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("+01:15").getId());"#);
    assert_eq!(out, vec!["+01:15"]);
}

#[test]
fn zone_id_rules_is_fixed_offset_utc() {
    let out =
        run_main(r#"System.out.println(java.time.ZoneId.of("UTC").getRules().isFixedOffset());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zone_id_rules_is_fixed_offset_paris_false() {
    let out = run_main(
        r#"System.out.println(java.time.ZoneId.of("Europe/Paris").getRules().isFixedOffset());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn zone_id_of_indian_mauritius() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("Indian/Mauritius").getId());"#);
    assert_eq!(out, vec!["Indian/Mauritius"]);
}

#[test]
fn zone_id_of_atlantic_reykjavik() {
    let out = run_main(r#"System.out.println(java.time.ZoneId.of("Atlantic/Reykjavik").getId());"#);
    assert_eq!(out, vec!["Atlantic/Reykjavik"]);
}

#[test]
fn zone_id_offset_at_instant_reykjavik() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2024-01-15T00:00:00Z"); System.out.println(java.time.ZoneId.of("Atlantic/Reykjavik").getRules().getOffset(i).getTotalSeconds());"#,
    );
    assert_eq!(out, vec!["0"]);
}
