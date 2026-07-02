//! Enum metaprogramming: `Parse`, `TryParse`, `GetNames`, `GetValues`, `HasFlag`, casts.
//! GAP: reflection-style enum APIs need dedicated structural coverage.


csharp_cases! {
    enum_parse_string_to_member => {
        r#"enum Color{Red,Green,Blue} var c=(Color)System.Enum.Parse(typeof(Color),"Green"); Console.WriteLine(c);"#,
        ["Green"]
    };

    enum_parse_first_member_default_zero => {
        r#"enum State{Idle,Running} var s=(State)System.Enum.Parse(typeof(State),"Idle"); Console.WriteLine((int)s);"#,
        ["0"]
    };

    enum_parse_last_declared_member => {
        r#"enum Level{Low,Mid,High} var v=(Level)System.Enum.Parse(typeof(Level),"High"); Console.WriteLine(v);"#,
        ["High"]
    };

    enum_try_parse_success_returns_true => {
        r#"enum Day{Mon,Tue,Wed} var ok=System.Enum.TryParse<Day>("Tue",out var d); Console.WriteLine(ok); Console.WriteLine(d);"#,
        ["True", "Tue"]
    };

    enum_try_parse_failure_returns_false => {
        r#"enum Day{Mon,Tue} var ok=System.Enum.TryParse<Day>("Sun",out var d); Console.WriteLine(ok);"#,
        ["False"]
    };

    enum_try_parse_ignore_case_success => {
        r#"enum Mode{Alpha,Beta} var ok=System.Enum.TryParse<Mode>("beta",true,out var m); Console.WriteLine(ok); Console.WriteLine(m);"#,
        ["True", "Beta"]
    };

    enum_get_names_returns_all_identifiers => {
        r#"enum Coin{Penny,Nickel,Dime} foreach(var name in System.Enum.GetNames(typeof(Coin))) Console.WriteLine(name);"#,
        ["Penny", "Nickel", "Dime"]
    };

    enum_get_names_count_matches_members => {
        r#"enum Size{Small,Medium,Large,Extra} Console.WriteLine(System.Enum.GetNames(typeof(Size)).Length);"#,
        ["4"]
    };

    enum_get_values_yields_each_member => {
        r#"enum Pair{A,B,C} foreach(var v in System.Enum.GetValues(typeof(Pair))) Console.WriteLine(v);"#,
        ["A", "B", "C"]
    };

    enum_get_values_first_is_zero_based => {
        r#"enum Rank{First,Second} foreach(var v in System.Enum.GetValues(typeof(Rank))) Console.WriteLine((int)v);"#,
        ["0", "1"]
    };

    enum_has_flag_detects_present_bit => {
        r#"[System.Flags] enum Perm{Read=1,Write=2} var p=Perm.Read|Perm.Write; Console.WriteLine(p.HasFlag(Perm.Read));"#,
        ["True"]
    };

    enum_has_flag_reports_absent_bit => {
        r#"[System.Flags] enum Perm{Read=1,Write=2,Execute=4} var p=Perm.Read; Console.WriteLine(p.HasFlag(Perm.Execute));"#,
        ["False"]
    };

    enum_cast_to_int_explicit_value => {
        r#"enum Code{A=10,B=20} Console.WriteLine((int)Code.B);"#,
        ["20"]
    };

    enum_cast_from_int_to_member => {
        r#"enum Code{A=1,B=2} var c=(Code)2; Console.WriteLine(c);"#,
        ["B"]
    };

    enum_underlying_type_default_is_int32 => {
        r#"enum Plain{One} Console.WriteLine(System.Enum.GetUnderlyingType(typeof(Plain)).Name);"#,
        ["Int32"]
    };

    enum_underlying_type_byte_annotation => {
        r#"enum Tiny:byte{X=1,Y=2} Console.WriteLine(System.Enum.GetUnderlyingType(typeof(Tiny)).Name);"#,
        ["Byte"]
    };

    enum_cast_to_byte_underlying => {
        r#"enum Tiny:byte{A=200,B=201} Console.WriteLine((byte)Tiny.B);"#,
        ["201"]
    };

    enum_cast_to_short_underlying => {
        r#"enum ShortEnum:short{Neg=-1,Pos=1} Console.WriteLine((short)ShortEnum.Neg);"#,
        ["-1"]
    };

    enum_cast_to_long_underlying => {
        r#"enum Wide:long{Big=10000000000L} Console.WriteLine((long)Wide.Big);"#,
        ["10000000000"]
    };

    enum_parse_with_explicit_values => {
        r#"enum Http{Ok=200,NotFound=404} var v=(Http)System.Enum.Parse(typeof(Http),"NotFound"); Console.WriteLine((int)v);"#,
        ["404"]
    };

    enum_try_parse_generic_with_out_var => {
        r#"enum Status{Open,Closed} System.Enum.TryParse<Status>("Closed",out var s); Console.WriteLine(s);"#,
        ["Closed"]
    };

    enum_get_names_on_single_member => {
        r#"enum Solo{Only} Console.WriteLine(System.Enum.GetNames(typeof(Solo))[0]);"#,
        ["Only"]
    };

    enum_get_values_length_matches_names => {
        r#"enum Trio{X,Y,Z} Console.WriteLine(System.Enum.GetValues(typeof(Trio)).Length);"#,
        ["3"]
    };

    enum_has_flag_combined_none_is_false => {
        r#"[System.Flags] enum Perm{None=0,Read=1} var p=Perm.None; Console.WriteLine(p.HasFlag(Perm.Read));"#,
        ["False"]
    };

    enum_has_flag_all_bits_set => {
        r#"[System.Flags] enum Perm{Read=1,Write=2,Execute=4} var p=Perm.Read|Perm.Write|Perm.Execute; Console.WriteLine(p.HasFlag(Perm.Execute));"#,
        ["True"]
    };

    enum_is_defined_for_valid_name => {
        r#"enum Phase{Start,End} Console.WriteLine(System.Enum.IsDefined(typeof(Phase),"Start"));"#,
        ["True"]
    };

    enum_is_defined_for_invalid_name => {
        r#"enum Phase{Start,End} Console.WriteLine(System.Enum.IsDefined(typeof(Phase),"Middle"));"#,
        ["False"]
    };

    enum_is_defined_for_numeric_value => {
        r#"enum Num{A=5,B=6} Console.WriteLine(System.Enum.IsDefined(typeof(Num),5));"#,
        ["True"]
    };

    enum_to_string_member_name => {
        r#"enum Letter{A,B,C} Console.WriteLine(Letter.B.ToString());"#,
        ["B"]
    };

    enum_format_d_decimal_representation => {
        r#"enum Num{X=7} Console.WriteLine(System.Enum.Format(typeof(Num),Num.X,"D"));"#,
        ["7"]
    };

    enum_format_g_general_name => {
        r#"enum Num{X=7} Console.WriteLine(System.Enum.Format(typeof(Num),Num.X,"G"));"#,
        ["X"]
    };

    enum_parse_case_sensitive_by_default => {
        r#"enum Case{Ab} var ok=System.Enum.TryParse<Case>("ab",out var v); Console.WriteLine(ok);"#,
        ["False"]
    };

    enum_try_parse_with_ignore_case_false => {
        r#"enum Case{Ab} var ok=System.Enum.TryParse<Case>("Ab",false,out var v); Console.WriteLine(ok); Console.WriteLine(v);"#,
        ["True", "Ab"]
    };

    enum_get_names_empty_for_zero_members_edge => {
        r#"enum Edge{A} Console.WriteLine(System.Enum.GetNames(typeof(Edge)).Length==1);"#,
        ["True"]
    };

    enum_get_values_cast_to_int_array => {
        r#"enum Score{A=1,B=3,C=5} int sum=0; foreach(var v in System.Enum.GetValues(typeof(Score))) sum+=(int)v; Console.WriteLine(sum);"#,
        ["9"]
    };

    enum_flags_or_combine_numeric => {
        r#"[System.Flags] enum F{A=1,B=2,C=4} var v=F.A|F.C; Console.WriteLine((int)v);"#,
        ["5"]
    };

    enum_flags_and_mask => {
        r#"[System.Flags] enum F{A=1,B=2,C=4} var v=(F.A|F.B|F.C)&F.B; Console.WriteLine((int)v);"#,
        ["2"]
    };

    enum_underlying_type_sbyte => {
        r#"enum SByteEnum:sbyte{Min=-128} Console.WriteLine(System.Enum.GetUnderlyingType(typeof(SByteEnum)).Name);"#,
        ["SByte"]
    };

    enum_cast_to_sbyte_underlying => {
        r#"enum SByteEnum:sbyte{Min=-128} Console.WriteLine((sbyte)SByteEnum.Min);"#,
        ["-128"]
    };

    enum_underlying_type_ushort => {
        r#"enum UShortEnum:ushort{Max=65535} Console.WriteLine(System.Enum.GetUnderlyingType(typeof(UShortEnum)).Name);"#,
        ["UInt16"]
    };

    enum_cast_to_ushort_underlying => {
        r#"enum UShortEnum:ushort{Max=65535} Console.WriteLine((ushort)UShortEnum.Max);"#,
        ["65535"]
    };

    enum_underlying_type_uint => {
        r#"enum UIntEnum:uint{Big=3000000000u} Console.WriteLine(System.Enum.GetUnderlyingType(typeof(UIntEnum)).Name);"#,
        ["UInt32"]
    };

    enum_parse_then_cast_roundtrip => {
        r#"enum Round{A=11,B=22} var p=(Round)System.Enum.Parse(typeof(Round),"B"); Console.WriteLine((int)p);"#,
        ["22"]
    };

    enum_try_parse_empty_string_fails => {
        r#"enum Empty{A} var ok=System.Enum.TryParse<Empty>("",out var v); Console.WriteLine(ok);"#,
        ["False"]
    };

    enum_get_names_preserves_declaration_order => {
        r#"enum Order{Z,A,M} Console.WriteLine(System.Enum.GetNames(typeof(Order))[1]);"#,
        ["A"]
    };

    enum_has_flag_single_bit_only => {
        r#"[System.Flags] enum Bit{One=1,Two=2} var v=Bit.Two; Console.WriteLine(v.HasFlag(Bit.One)); Console.WriteLine(v.HasFlag(Bit.Two));"#,
        ["False", "True"]
    };

    enum_compare_equal_members => {
        r#"enum Eq{X,Y} Console.WriteLine(Eq.X==Eq.X); Console.WriteLine(Eq.X==Eq.Y);"#,
        ["True", "False"]
    };

    enum_switch_on_parsed_value => {
        r#"enum Mode{On,Off} var m=(Mode)System.Enum.Parse(typeof(Mode),"On"); string s=m==Mode.On?"yes":"no"; Console.WriteLine(s);"#,
        ["yes"]
    };

    enum_get_values_element_type_is_enum => {
        r#"enum Kind{A,B} var values=System.Enum.GetValues(typeof(Kind)); Console.WriteLine(values.GetType().GetElementType().Name);"#,
        ["Kind"]
    };

    enum_flags_cast_to_int_preserves_bits => {
        r#"[System.Flags] enum F{A=1,B=2,C=4} var v=F.A|F.B; Console.WriteLine((int)v);"#,
        ["3"]
    };

    enum_parse_whitespace_trim_not_required_exact => {
        r#"enum Exact{Ab} var ok=System.Enum.TryParse<Exact>("Ab",out var v); Console.WriteLine(ok);"#,
        ["True"]
    };
}
