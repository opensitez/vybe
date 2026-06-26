//! Property patterns: `{ Prop: val }`, nested `{ Outer: { Inner: val } }`, and switch arms.

csharp_cases! {
    is_property_pattern_matches_exact_int_field => {
        r#"class Box { public int Value; } object o=new Box{Value=10}; Console.WriteLine(o is Box{Value:10});"#,
        ["True"]
    };

    is_property_pattern_rejects_wrong_int_field => {
        r#"class Box { public int Value; } object o=new Box{Value=10}; Console.WriteLine(o is Box{Value:11});"#,
        ["False"]
    };

    is_property_pattern_var_capture_reads_field => {
        r#"class Box { public int Value; } object o=new Box{Value=25}; if(o is Box{Value:var v}) Console.WriteLine(v);"#,
        ["25"]
    };

    is_property_pattern_bool_true_literal => {
        r#"class Flag { public bool On; } object o=new Flag{On=true}; Console.WriteLine(o is Flag{On:true});"#,
        ["True"]
    };

    is_property_pattern_bool_false_literal => {
        r#"class Flag { public bool On; } object o=new Flag{On=false}; Console.WriteLine(o is Flag{On:false});"#,
        ["True"]
    };

    is_property_pattern_string_literal_match => {
        r#"class Tag { public string Name; } object o=new Tag{Name="alpha"}; Console.WriteLine(o is Tag{Name:"alpha"});"#,
        ["True"]
    };

    is_property_pattern_string_literal_mismatch => {
        r#"class Tag { public string Name; } object o=new Tag{Name="alpha"}; Console.WriteLine(o is Tag{Name:"beta"});"#,
        ["False"]
    };

    is_property_pattern_two_fields_both_required => {
        r#"class Pair { public int A; public int B; } object o=new Pair{A=2,B=3}; Console.WriteLine(o is Pair{A:2,B:3});"#,
        ["True"]
    };

    is_property_pattern_two_fields_one_wrong => {
        r#"class Pair { public int A; public int B; } object o=new Pair{A=2,B=3}; Console.WriteLine(o is Pair{A:2,B:4});"#,
        ["False"]
    };

    is_property_pattern_relational_greater_on_field => {
        r#"class Score { public int Points; } object o=new Score{Points=95}; Console.WriteLine(o is Score{Points:>90});"#,
        ["True"]
    };

    is_property_pattern_relational_less_on_field => {
        r#"class Score { public int Points; } object o=new Score{Points=40}; Console.WriteLine(o is Score{Points:<50});"#,
        ["True"]
    };

    is_property_pattern_relational_greater_equal_boundary => {
        r#"class Level { public int Tier; } object o=new Level{Tier=3}; Console.WriteLine(o is Level{Tier:>=3});"#,
        ["True"]
    };

    is_property_pattern_relational_less_equal_boundary => {
        r#"class Level { public int Tier; } object o=new Level{Tier=3}; Console.WriteLine(o is Level{Tier:<=3});"#,
        ["True"]
    };

    is_property_pattern_and_two_relational_fields => {
        r#"class Range { public int Lo; public int Hi; } object o=new Range{Lo=5,Hi=15}; Console.WriteLine(o is Range{Lo:>0,Hi:<20});"#,
        ["True"]
    };

    nested_property_pattern_matches_inner_field => {
        r#"class Inner { public int N; } class Outer { public Inner Child; } object o=new Outer{Child=new Inner{N=7}}; if(o is Outer{Child:{N:7}}) Console.WriteLine("ok");"#,
        ["ok"]
    };

    nested_property_pattern_var_capture_inner => {
        r#"class Inner { public int N; } class Outer { public Inner Child; } object o=new Outer{Child=new Inner{N=9}}; if(o is Outer{Child:{N:var n}}) Console.WriteLine(n);"#,
        ["9"]
    };

    nested_property_pattern_string_on_inner => {
        r#"class Address { public string City; } class Person { public Address Home; } object p=new Person{Home=new Address{City="Paris"}}; Console.WriteLine(p is Person{Home:{City:"Paris"}});"#,
        ["True"]
    };

    nested_property_pattern_rejects_wrong_inner => {
        r#"class Address { public string City; } class Person { public Address Home; } object p=new Person{Home=new Address{City="Paris"}}; Console.WriteLine(p is Person{Home:{City:"London"}});"#,
        ["False"]
    };

    triple_nested_property_pattern_matches_leaf => {
        r#"class Leaf { public int V; } class Mid { public Leaf L; } class Root { public Mid M; } object o=new Root{M=new Mid{L=new Leaf{V=4}}}; if(o is Root{M:{L:{V:4}}}) Console.WriteLine("deep");"#,
        ["deep"]
    };

    switch_expression_property_pattern_big_paid_order => {
        r#"class Order { public int Amount; public bool Paid; } string Label(object o)=>o switch{Order{Paid:true,Amount:>50}=>"big-paid",Order{Paid:true}=>"paid",_=>"open"}; Console.WriteLine(Label(new Order{Amount=100,Paid=true}));"#,
        ["big-paid"]
    };

    switch_expression_property_pattern_small_paid_order => {
        r#"class Order { public int Amount; public bool Paid; } string Label(object o)=>o switch{Order{Paid:true,Amount:>50}=>"big-paid",Order{Paid:true}=>"paid",_=>"open"}; Console.WriteLine(Label(new Order{Amount=10,Paid=true}));"#,
        ["paid"]
    };

    switch_expression_property_pattern_open_order => {
        r#"class Order { public int Amount; public bool Paid; } string Label(object o)=>o switch{Order{Paid:true,Amount:>50}=>"big-paid",Order{Paid:true}=>"paid",_=>"open"}; Console.WriteLine(Label(new Order{Amount=10,Paid=false}));"#,
        ["open"]
    };

    switch_expression_property_pattern_capture_amount => {
        r#"class Wallet { public int Balance; } int Read(object o)=>o switch{Wallet{Balance:var b}=>b,_=>-1}; Console.WriteLine(Read(new Wallet{Balance=42}));"#,
        ["42"]
    };

    switch_expression_nested_property_city_label => {
        r#"class Address { public string City; } class Person { public Address Addr; } string Where(object p)=>p switch{Person{Addr:{City:"NYC"}}=>"metro",_=>"other"}; Console.WriteLine(Where(new Person{Addr=new Address{City="NYC"}}));"#,
        ["metro"]
    };

    switch_statement_property_pattern_case => {
        r#"class Node { public int Id; } object o=new Node{Id=5}; string tag=""; switch(o){case Node{Id:5}:tag="match";break;default:tag="miss";break;} Console.WriteLine(tag);"#,
        ["match"]
    };

    switch_statement_property_pattern_case_with_capture => {
        r#"class Node { public int Id; } object o=new Node{Id=12}; string tag=""; switch(o){case Node{Id:var id}:tag=id.ToString();break;default:tag="0";break;} Console.WriteLine(tag);"#,
        ["12"]
    };

    record_property_pattern_positional_and_named => {
        r#"record Point(int X,int Y); object o=new Point(1,2); Console.WriteLine(o is Point{X:1,Y:2});"#,
        ["True"]
    };

    record_property_pattern_capture_y => {
        r#"record Point(int X,int Y); object o=new Point(3,8); if(o is Point{X:3,Y:var y}) Console.WriteLine(y);"#,
        ["8"]
    };

    is_property_pattern_partial_single_field_ignores_rest => {
        r#"class Wide { public int A; public int B; public int C; } object o=new Wide{A=1,B=2,C=3}; Console.WriteLine(o is Wide{A:1});"#,
        ["True"]
    };

    is_property_pattern_enum_field_match => {
        r#"enum Color { Red, Green } class Paint { public Color Hue; } object o=new Paint{Hue=Color.Green}; Console.WriteLine(o is Paint{Hue:Color.Green});"#,
        ["True"]
    };

    is_property_pattern_nullable_int_has_value => {
        r#"class Holder { public int? Slot; } object o=new Holder{Slot=6}; Console.WriteLine(o is Holder{Slot:6});"#,
        ["True"]
    };

    is_property_pattern_nullable_int_null_arm => {
        r#"class Holder { public int? Slot; } object o=new Holder{Slot=null}; Console.WriteLine(o is Holder{Slot:null});"#,
        ["True"]
    };

    is_property_pattern_char_field_literal => {
        r#"class Glyph { public char Ch; } object o=new Glyph{Ch='Z'}; Console.WriteLine(o is Glyph{Ch:'Z'});"#,
        ["True"]
    };

    is_property_pattern_double_field_literal => {
        r#"class Measure { public double M; } object o=new Measure{M=2.5}; Console.WriteLine(o is Measure{M:2.5});"#,
        ["True"]
    };

    is_property_pattern_long_field_literal => {
        r#"class Span { public long L; } object o=new Span{L=1000L}; Console.WriteLine(o is Span{L:1000L});"#,
        ["True"]
    };

    is_property_pattern_not_inverts_match => {
        r#"class Box { public int V; } object o=new Box{V=1}; Console.WriteLine(o is not Box{V:2});"#,
        ["True"]
    };

    is_property_pattern_with_when_on_captured_var => {
        r#"class Pair { public int A; public int B; } object o=new Pair{A=4,B=9}; if(o is Pair{A:var a,B:var b} when a<b) Console.WriteLine(b-a);"#,
        ["5"]
    };

    switch_expression_property_when_guard_on_fields => {
        r#"class Pair { public int A; public int B; } string Sign(object o)=>o switch{Pair{A:var x,B:var y} when x==y=>"eq",Pair{A:var x,B:var y}=>"neq",_=>"?"}; Console.WriteLine(Sign(new Pair{A=3,B:3}));"#,
        ["eq"]
    };

    inheritance_property_pattern_on_derived_type => {
        r#"class Animal { public string Kind; } class Dog : Animal { public int Legs; } object o=new Dog{Kind="pet",Legs=4}; Console.WriteLine(o is Dog{Legs:4,Kind:"pet"});"#,
        ["True"]
    };

    struct_property_pattern_on_value_type => {
        r#"struct Vec2 { public int X; public int Y; } object o=new Vec2{X=2,Y=3}; Console.WriteLine(o is Vec2{X:2,Y:3});"#,
        ["True"]
    };

    nested_property_pattern_two_levels_with_relational => {
        r#"class Inner { public int N; } class Outer { public Inner I; } object o=new Outer{I=new Inner{N=50}}; Console.WriteLine(o is Outer{I:{N:>40}});"#,
        ["True"]
    };

    switch_expression_property_relational_amount_tier => {
        r#"class Bill { public int Amount; } string Tier(object o)=>o switch{Bill{Amount:>=100}=>"gold",Bill{Amount:>=50}=>"silver",_=>"bronze"}; Console.WriteLine(Tier(new Bill{Amount=75}));"#,
        ["silver"]
    };

    switch_expression_property_relational_amount_gold => {
        r#"class Bill { public int Amount; } string Tier(object o)=>o switch{Bill{Amount:>=100}=>"gold",Bill{Amount:>=50}=>"silver",_=>"bronze"}; Console.WriteLine(Tier(new Bill{Amount=120}));"#,
        ["gold"]
    };

    is_property_pattern_string_var_capture => {
        r#"class Label { public string Text; } object o=new Label{Text="go"}; if(o is Label{Text:var t}) Console.WriteLine(t.ToUpper());"#,
        ["GO"]
    };

    nested_property_pattern_three_string_fields => {
        r#"class Street { public string Name; } class Addr { public Street S; } class Person { public Addr A; } object p=new Person{A=new Addr{S=new Street{Name="Main"}}}; Console.WriteLine(p is Person{A:{S:{Name:"Main"}}});"#,
        ["True"]
    };

    is_property_pattern_and_relational_on_same_type => {
        r#"class Temp { public int C; } object o=new Temp{C=22}; Console.WriteLine(o is Temp{C:>=20 and <=25});"#,
        ["True"]
    };

    switch_expression_property_nested_capture_sum => {
        r#"class Inner { public int A; public int B; } class Wrap { public Inner Data; } int Sum(object o)=>o switch{Wrap{Data:{A:var a,B:var b}}=>a+b,_=>0}; Console.WriteLine(Sum(new Wrap{Data=new Inner{A=6,B=7}}));"#,
        ["13"]
    };

    is_property_pattern_zero_field_value => {
        r#"class Zero { public int Z; } object o=new Zero{Z=0}; Console.WriteLine(o is Zero{Z:0});"#,
        ["True"]
    };

    is_property_pattern_negative_field_value => {
        r#"class Delta { public int D; } object o=new Delta{D=-5}; Console.WriteLine(o is Delta{D:-5});"#,
        ["True"]
    };

    switch_expression_property_default_after_specific_arms => {
        r#"class Token { public string Kind; } string Name(object o)=>o switch{Token{Kind:"add"}=>"plus",Token{Kind:"sub"}=>"minus",_=>"other"}; Console.WriteLine(Name(new Token{Kind="mul"}));"#,
        ["other"]
    };

    is_property_pattern_multiple_relational_and_combo => {
        r#"class Band { public int Lo; public int Hi; } object o=new Band{Lo=10,Hi=20}; Console.WriteLine(o is Band{Lo:>=10 and <=10,Hi:>=20 and <=20});"#,
        ["True"]
    };

    switch_expression_property_when_false_falls_through => {
        r#"class Item { public int Q; } string Flag(object o)=>o switch{Item{Q:var q} when q>10=>"big",Item{Q:var q}=>"small",_=>"?"}; Console.WriteLine(Flag(new Item{Q=3}));"#,
        ["small"]
    };

    is_property_pattern_byte_field_literal => {
        r#"class Port { public byte P; } object o=new Port{P=80}; Console.WriteLine(o is Port{P:80});"#,
        ["True"]
    };

    is_property_pattern_float_field_literal => {
        r#"class Rate { public float R; } object o=new Rate{R=1.5f}; Console.WriteLine(o is Rate{R:1.5f});"#,
        ["True"]
    };

    switch_expression_property_or_literal_kind_arms => {
        r#"class Msg { public string Kind; } string Label(object o)=>o switch{Msg{Kind:"err" or "fail"}=>"bad",_=>"ok"}; Console.WriteLine(Label(new Msg{Kind="fail"}));"#,
        ["bad"]
    };
}
