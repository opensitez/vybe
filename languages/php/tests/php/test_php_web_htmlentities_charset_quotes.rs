use super::helpers::run_prints;

#[test]
fn test_htmlentities_ent_quotes_mode() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "<p> 'single' & \"double\" </p>";
echo htmlentities($str, ENT_QUOTES, 'UTF-8'), "\n";
"#
        ),
        vec!["&lt;p&gt; &#039;single&#039; &amp; &quot;double&quot; &lt;/p&gt;"]
    );
}
