//! Relational patterns: `>`, `<`, `>=`, `<=`, combined with `and` / `or` in `is` and switch expressions.

csharp_cases! {
    is_relational_greater_than_matches_above_threshold => {
        r#"int n=5; Console.WriteLine(n is >3);"#,
        ["True"]
    };

    is_relational_greater_than_rejects_equal_value => {
        r#"int n=3; Console.WriteLine(n is >3);"#,
        ["False"]
    };

    is_relational_less_than_matches_below_threshold => {
        r#"int n=2; Console.WriteLine(n is <5);"#,
        ["True"]
    };

    is_relational_less_than_rejects_equal_value => {
        r#"int n=5; Console.WriteLine(n is <5);"#,
        ["False"]
    };

    is_relational_greater_equal_matches_boundary => {
        r#"int n=10; Console.WriteLine(n is >=10);"#,
        ["True"]
    };

    is_relational_greater_equal_accepts_above => {
        r#"int n=11; Console.WriteLine(n is >=10);"#,
        ["True"]
    };

    is_relational_less_equal_matches_boundary => {
        r#"int n=10; Console.WriteLine(n is <=10);"#,
        ["True"]
    };

    is_relational_less_equal_accepts_below => {
        r#"int n=9; Console.WriteLine(n is <=10);"#,
        ["True"]
    };

    is_relational_negative_value_less_than_zero => {
        r#"int n=-1; Console.WriteLine(n is <0);"#,
        ["True"]
    };

    is_relational_negative_value_greater_than_minus_ten => {
        r#"int n=-3; Console.WriteLine(n is >-10);"#,
        ["True"]
    };

    is_relational_and_range_matches_interior => {
        r#"int n=15; Console.WriteLine(n is >10 and <20);"#,
        ["True"]
    };

    is_relational_and_range_rejects_lower_bound => {
        r#"int n=10; Console.WriteLine(n is >10 and <20);"#,
        ["False"]
    };

    is_relational_and_range_rejects_upper_bound => {
        r#"int n=20; Console.WriteLine(n is >10 and <20);"#,
        ["False"]
    };

    is_relational_and_closed_interval_matches_edge => {
        r#"int n=80; Console.WriteLine(n is >=80 and <=89);"#,
        ["True"]
    };

    is_relational_or_matches_first_branch => {
        r#"int n=3; Console.WriteLine(n is <0 or >10);"#,
        ["False"]
    };

    is_relational_or_matches_second_branch => {
        r#"int n=15; Console.WriteLine(n is <0 or >10);"#,
        ["True"]
    };

    is_relational_or_matches_either_negative => {
        r#"int n=-2; Console.WriteLine(n is <0 or >10);"#,
        ["True"]
    };

    switch_expression_relational_greater_arm_grade_a => {
        r#"int score=95; Console.WriteLine(score switch{>=90=>"A",_=>"B"});"#,
        ["A"]
    };

    switch_expression_relational_greater_arm_grade_b => {
        r#"int score=85; Console.WriteLine(score switch{>=90=>"A",>=80=>"B",_=>"C"});"#,
        ["B"]
    };

    switch_expression_relational_less_arm_negative => {
        r#"int n=-4; Console.WriteLine(n switch{<0=>"neg",0=>"zero",_=>"pos"});"#,
        ["neg"]
    };

    switch_expression_relational_zero_constant_arm => {
        r#"int n=0; Console.WriteLine(n switch{<0=>"neg",0=>"zero",>0=>"pos"});"#,
        ["zero"]
    };

    switch_expression_relational_positive_arm => {
        r#"int n=7; Console.WriteLine(n switch{<0=>"neg",0=>"zero",>0=>"pos"});"#,
        ["pos"]
    };

    switch_expression_relational_and_band_b => {
        r#"int n=75; Console.WriteLine(n switch{>=90=>"A",>=70 and <90=>"B",_=>"F"});"#,
        ["B"]
    };

    switch_expression_relational_and_band_c => {
        r#"int n=55; Console.WriteLine(n switch{>=90=>"A",>=70 and <90=>"B",>=50 and <70=>"C",_=>"F"});"#,
        ["C"]
    };

    switch_expression_relational_and_band_f => {
        r#"int n=30; Console.WriteLine(n switch{>=90=>"A",>=70 and <90=>"B",>=50 and <70=>"C",_=>"F"});"#,
        ["F"]
    };

    switch_expression_relational_less_equal_upper_cap => {
        r#"int n=100; Console.WriteLine(n switch{<=100=>"ok",_=>"high"});"#,
        ["ok"]
    };

    switch_expression_relational_greater_equal_lower_cap => {
        r#"int n=1; Console.WriteLine(n switch{>=1=>"ok",_=>"low"});"#,
        ["ok"]
    };

    is_relational_on_byte_value => {
        r#"byte b=200; Console.WriteLine(b is >100);"#,
        ["True"]
    };

    is_relational_on_long_value => {
        r#"long x=5000L; Console.WriteLine(x is >=1000L);"#,
        ["True"]
    };

    is_relational_on_double_value => {
        r#"double d=3.14; Console.WriteLine(d is >3.0);"#,
        ["True"]
    };

    is_relational_not_inverts_greater_match => {
        r#"int n=2; Console.WriteLine(n is not >5);"#,
        ["True"]
    };

    is_relational_not_inverts_failed_less => {
        r#"int n=8; Console.WriteLine(n is not <5);"#,
        ["True"]
    };

    switch_expression_relational_chained_thresholds => {
        r#"int v=42; Console.WriteLine(v switch{<10=>"xs",<100=>"md",_=>"lg"});"#,
        ["md"]
    };

    switch_expression_relational_chained_large => {
        r#"int v=150; Console.WriteLine(v switch{<10=>"xs",<100=>"md",_=>"lg"});"#,
        ["lg"]
    };

    is_relational_and_with_or_inside => {
        r#"int n=12; Console.WriteLine(n is (>10 and <20) or >100);"#,
        ["True"]
    };

    is_relational_or_with_and_groups => {
        r#"int n=5; Console.WriteLine(n is <0 or (>=5 and <=5));"#,
        ["True"]
    };

    switch_expression_relational_with_when_guard => {
        r#"int n=8; Console.WriteLine(n switch{int x when x>5 and x<10=>"mid",_=>"other"});"#,
        ["mid"]
    };

    switch_expression_relational_when_guard_fallback => {
        r#"int n=2; Console.WriteLine(n switch{int x when x>5 and x<10=>"mid",_=>"other"});"#,
        ["other"]
    };

    is_relational_zero_greater_equal_zero => {
        r#"int n=0; Console.WriteLine(n is >=0);"#,
        ["True"]
    };

    is_relational_zero_less_equal_zero => {
        r#"int n=0; Console.WriteLine(n is <=0);"#,
        ["True"]
    };

    switch_expression_relational_enum_underlying_int => {
        r#"enum Tier { Low=1, Mid=5, High=10 } var t=Tier.Mid; Console.WriteLine(t switch{>=Tier.Mid=>"up",_=>"down"});"#,
        ["up"]
    };

    is_relational_object_boxed_int_greater => {
        r#"object o=9; Console.WriteLine(o is int n and >5);"#,
        ["True"]
    };

    is_relational_object_boxed_int_range => {
        r#"object o=7; Console.WriteLine(o is int n and >=7 and <=7);"#,
        ["True"]
    };

    switch_expression_relational_returns_int_from_arms => {
        r#"int n=6; Console.WriteLine(n switch{>10=>20,>5=>10,_=>0});"#,
        ["10"]
    };

    switch_expression_relational_returns_int_default => {
        r#"int n=3; Console.WriteLine(n switch{>10=>20,>5=>10,_=>0});"#,
        ["0"]
    };

    is_relational_and_three_part_window => {
        r#"int n=50; Console.WriteLine(n is >=40 and <=60 and !=55);"#,
        ["True"]
    };

    is_relational_and_rejects_excluded_middle => {
        r#"int n=55; Console.WriteLine(n is >=40 and <=60 and !=55);"#,
        ["False"]
    };

    switch_expression_relational_or_literal_combo => {
        r#"int code=404; Console.WriteLine(code switch{200=>"ok",404 or 500=>"err",_=>"?"});"#,
        ["err"]
    };

    switch_expression_relational_or_literal_second => {
        r#"int code=500; Console.WriteLine(code switch{200=>"ok",404 or 500=>"err",_=>"?"});"#,
        ["err"]
    };

    is_relational_less_equal_negative_boundary => {
        r#"int n=-100; Console.WriteLine(n is <=-50);"#,
        ["True"]
    };

    switch_expression_relational_nested_selector => {
        r#"int a=2,b=3; Console.WriteLine((a+b) switch{>4=>"big",_=>"small"});"#,
        ["big"]
    };
}
