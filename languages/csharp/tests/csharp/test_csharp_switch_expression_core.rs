//! Switch expressions (`expr switch { ... }`): value yields, nesting, and `when` guards — not switch statements.

csharp_cases! {
    switch_expr_assign_to_local_int_literal_arm => {
        r#"var code=2; var label=code switch{1=>"one",2=>"two",_=>"many"}; Console.WriteLine(label);"#,
        ["two"]
    };

    switch_expr_assign_to_local_default_arm => {
        r#"var code=9; var label=code switch{1=>"one",2=>"two",_=>"many"}; Console.WriteLine(label);"#,
        ["many"]
    };

    switch_expr_return_from_local_function => {
        r#"string Word(int n)=>n switch{1=>"a",2=>"b",_=>"z"}; Console.WriteLine(Word(2));"#,
        ["b"]
    };

    switch_expr_nested_arm_switch_on_inner_value => {
        r#"string Outer(int n)=>n switch{1=>1 switch{1=>"one-one",_=>"one-other"},2=>"two",_=>"rest"}; Console.WriteLine(Outer(1));"#,
        ["one-one"]
    };

    switch_expr_nested_arm_switch_default_inner => {
        r#"string Outer(int n)=>n switch{1=>5 switch{5=>"five",_=>"not-five"},_=>"rest"}; Console.WriteLine(Outer(1));"#,
        ["not-five"]
    };

    switch_expr_double_nested_in_arm_result => {
        r#"int Pick(int a,int b)=>a switch{1=>b switch{2=>10,3=>20,_=>0},_=>-1}; Console.WriteLine(Pick(1,3));"#,
        ["20"]
    };

    switch_expr_triple_nested_selector_chain => {
        r#"int Depth(int n)=>n switch{1=>2 switch{2=>3 switch{3=>9,_=>0},_=>0},_=>0}; Console.WriteLine(Depth(1));"#,
        ["0"]
    };

    switch_expr_when_guard_matches_modulo_even => {
        r#"var x=12; Console.WriteLine(x switch{int n when n%2==0=>"even",int n=>"odd"});"#,
        ["even"]
    };

    switch_expr_when_guard_matches_modulo_odd => {
        r#"var x=11; Console.WriteLine(x switch{int n when n%2==0=>"even",int n=>"odd"});"#,
        ["odd"]
    };

    switch_expr_when_guard_string_length_four => {
        r#"var s="tool"; Console.WriteLine(s switch{string t when t.Length==4=>"len4",string t=>t.Length.ToString(),_=>"0"});"#,
        ["len4"]
    };

    switch_expr_when_guard_string_length_other => {
        r#"var s="hi"; Console.WriteLine(s switch{string t when t.Length==4=>"len4",string t=>t.Length.ToString(),_=>"0"});"#,
        ["2"]
    };

    switch_expr_when_guard_false_skips_to_next_arm => {
        r#"var x=3; Console.WriteLine(x switch{int n when n>10=>"big",int n when n>1=>"mid",_=>"small"});"#,
        ["mid"]
    };

    switch_expr_when_guard_first_true_stops => {
        r#"var x=15; Console.WriteLine(x switch{int n when n>10=>"big",int n when n>1=>"mid",_=>"small"});"#,
        ["big"]
    };

    switch_expr_when_on_type_pattern_int_positive => {
        r#"object o=8; Console.WriteLine(o switch{int n when n>0=>"pos",_=>"other"});"#,
        ["pos"]
    };

    switch_expr_when_on_type_pattern_int_non_positive => {
        r#"object o=-1; Console.WriteLine(o switch{int n when n>0=>"pos",_=>"other"});"#,
        ["other"]
    };

    switch_expr_type_pattern_string_upper => {
        r#"object o="abc"; Console.WriteLine(o switch{string s=>s.ToUpper(),_=>"?"});"#,
        ["ABC"]
    };

    switch_expr_type_pattern_int_increment => {
        r#"object o=6; Console.WriteLine(o switch{int n=>(n+1).ToString(),_=>"?"});"#,
        ["7"]
    };

    switch_expr_discard_only_arm => {
        r#"var x=99; Console.WriteLine(x switch{_=>"always"});"#,
        ["always"]
    };

    switch_expr_bool_true_arm => {
        r#"bool flag=true; Console.WriteLine(flag switch{true=>"yes",false=>"no"});"#,
        ["yes"]
    };

    switch_expr_bool_false_arm => {
        r#"bool flag=false; Console.WriteLine(flag switch{true=>"yes",false=>"no"});"#,
        ["no"]
    };

    switch_expr_string_multi_case_second => {
        r#"var key="go"; Console.WriteLine(key switch{"stop"=>"S","go"=>"G",_=>"?"});"#,
        ["G"]
    };

    switch_expr_string_multi_case_default => {
        r#"var key="run"; Console.WriteLine(key switch{"stop"=>"S","go"=>"G",_=>"?"});"#,
        ["?"]
    };

    switch_expr_double_literal_arm => {
        r#"double d=2.5; Console.WriteLine(d switch{2.5=>"half",_=>"other"});"#,
        ["half"]
    };

    switch_expr_nullable_int_null_arm => {
        r#"int? v=null; Console.WriteLine(v switch{null=>"nil",_=>"val"});"#,
        ["nil"]
    };

    switch_expr_nullable_int_value_arm => {
        r#"int? v=7; Console.WriteLine(v switch{null=>"nil",_=>"val"});"#,
        ["val"]
    };

    switch_expr_enum_symbolic_arm => {
        r#"enum Mode { Off, On } var m=Mode.On; Console.WriteLine(m switch{Mode.Off=>"0",Mode.On=>"1",_=>"?"});"#,
        ["1"]
    };

    switch_expr_nested_in_addition_expression => {
        r#"var a=1,b=2; Console.WriteLine((a switch{1=>10,_=>0})+(b switch{2=>20,_=>0}));"#,
        ["30"]
    };

    switch_expr_as_method_argument => {
        r#"void Show(string s){Console.WriteLine(s);} Show(3 switch{3=>"three",_=>"other"});"#,
        ["three"]
    };

    switch_expr_in_string_concatenation => {
        r#"var n=4; Console.WriteLine("v="+(n switch{4=>"four",_=>"?"}));"#,
        ["v=four"]
    };

    switch_expr_lambda_body_returns_switch => {
        r#"var fn=(int x)=>x switch{0=>"z",_=>"nz"}; Console.WriteLine(fn(0));"#,
        ["z"]
    };

    switch_expr_parenthesized_arm_expression => {
        r#"var n=2; Console.WriteLine(n switch{2=>(3+4),_=>0});"#,
        ["7"]
    };

    switch_expr_arm_calls_method => {
        r#"int Double(int x)=>x*2; Console.WriteLine(5 switch{5=>Double(5),_=>0});"#,
        ["10"]
    };

    switch_expr_selector_is_method_call => {
        r#"int Twice(int x)=>x*2; Console.WriteLine(Twice(3) switch{6=>"ok",_=>"no"});"#,
        ["ok"]
    };

    switch_expr_relational_arm_less_than => {
        r#"var n=2; Console.WriteLine(n switch{<5=>"low",_=>"high"});"#,
        ["low"]
    };

    switch_expr_relational_arm_greater_equal => {
        r#"var n=10; Console.WriteLine(n switch{>=10=>"ten+",_=>"low"});"#,
        ["ten+"]
    };

    switch_expr_relational_and_arm_band => {
        r#"var n=85; Console.WriteLine(n switch{>=80 and <90=>"B",_=>"other"});"#,
        ["B"]
    };

    switch_expr_tuple_pattern_origin => {
        r#"var p=(0,0); Console.WriteLine(p switch{(0,0)=>"origin",_=>"away"});"#,
        ["origin"]
    };

    switch_expr_tuple_pattern_discard_axis => {
        r#"var p=(3,0); Console.WriteLine(p switch{(0,0)=>"origin",(_,0)=>"x",_=>"away"});"#,
        ["x"]
    };

    switch_expr_nested_switch_as_arm_value => {
        r#"var tier=2; Console.WriteLine(tier switch{1=>"a",2=>(3 switch{3=>"inner",_=>"outer"}),_=>"?"});"#,
        ["outer"]
    };

    switch_expr_when_and_relational_combo => {
        r#"var n=18; Console.WriteLine(n switch{int x when x>=18 and x<21=>"adult",_=>"minor"});"#,
        ["adult"]
    };

    switch_expr_when_and_relational_minor => {
        r#"var n=12; Console.WriteLine(n switch{int x when x>=18 and x<21=>"adult",_=>"minor"});"#,
        ["minor"]
    };

    switch_expr_or_constant_arms_weekend => {
        r#"var day="Sunday"; Console.WriteLine(day switch{"Saturday" or "Sunday"=>"off",_=>"work"});"#,
        ["off"]
    };

    switch_expr_or_constant_arms_weekday => {
        r#"var day="Monday"; Console.WriteLine(day switch{"Saturday" or "Sunday"=>"off",_=>"work"});"#,
        ["work"]
    };

    switch_expr_interpolated_result_arm => {
        r#"var score=88; Console.WriteLine(score switch{>=90=>$"A:{score}",>=80=>$"B:{score}",_=>$"C:{score}"});"#,
        ["B:88"]
    };

    switch_expr_array_initializer_element => {
        r#"var codes=new string[]{1 switch{1=>"one",_=>"?"}}; Console.WriteLine(codes[0]);"#,
        ["one"]
    };

    switch_expr_object_boxed_switch_type => {
        r#"object o=12L; Console.WriteLine(o switch{long l=>l.ToString(),int i=>i.ToString(),_=>"?"});"#,
        ["12"]
    };

    switch_expr_negative_int_literal_arm => {
        r#"var n=-2; Console.WriteLine(n switch{-2=>"neg-two",_=>"other"});"#,
        ["neg-two"]
    };

    switch_expr_zero_literal_explicit_arm => {
        r#"var n=0; Console.WriteLine(n switch{0=>"zero",_=>"nz"});"#,
        ["zero"]
    };

    switch_expr_char_literal_arm => {
        r#"char c='q'; Console.WriteLine(c switch{'q'=>"que",_=>"other"});"#,
        ["que"]
    };

    switch_expr_nested_when_on_outer_and_inner => {
        r#"var pair=(2,4); Console.WriteLine(pair switch{(var a,var b) when a<b=>"asc",(var a,var b)=>"desc",_=>"?"});"#,
        ["asc"]
    };

    switch_expr_when_on_outer_and_inner_desc => {
        r#"var pair=(5,2); Console.WriteLine(pair switch{(var a,var b) when a<b=>"asc",(var a,var b)=>"desc",_=>"?"});"#,
        ["desc"]
    };

    switch_expr_result_used_in_comparison => {
        r#"var n=3; Console.WriteLine((n switch{3=>10,_=>0})==10);"#,
        ["True"]
    };

    switch_expr_multiple_commas_trailing_ok => {
        r#"var n=1; Console.WriteLine(n switch{1=>"one",2=>"two",_=>"many" });"#,
        ["one"]
    };

    switch_expr_deep_nested_arm_three_levels => {
        r#"string L(int n)=>n switch{1=>"a",2=>2 switch{2=>"b",3=>"c",_=>"d"},_=>"z"}; Console.WriteLine(L(2));"#,
        ["b"]
    };

    switch_expr_var_pattern_binds_any_value => {
        r#"object o=42; Console.WriteLine(o switch{var x when x is int n and n>10=>"big",_=>"other"});"#,
        ["big"]
    };
}
