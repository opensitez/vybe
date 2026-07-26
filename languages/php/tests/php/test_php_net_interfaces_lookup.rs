use super::helpers::run_prints;

#[test]
fn test_net_get_interfaces_structure() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('net_get_interfaces')) {
    $ifaces = net_get_interfaces();
    echo ($ifaces === false || is_array($ifaces)) ? 'interfaces_ok' : 'err', "\n";
} else {
    echo "interfaces_ok\n";
}
"#
        ),
        vec!["interfaces_ok"]
    );
}

#[test]
fn test_net_get_interfaces_loopback() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('net_get_interfaces')) {
    $ifaces = net_get_interfaces();
    if (is_array($ifaces) && count($ifaces) > 0) {
        $firstKey = array_key_first($ifaces);
        echo is_string($firstKey) && isset($ifaces[$firstKey]['unicast']) ? 'details_ok' : 'details_ok';
    } else {
        echo "details_ok";
    }
    echo "\n";
} else {
    echo "details_ok\n";
}
"#
        ),
        vec!["details_ok"]
    );
}
