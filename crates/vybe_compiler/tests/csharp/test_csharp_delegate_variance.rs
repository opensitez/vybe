//! Delegate variance: `Func` covariance on return, `Action` contravariance on params.
//! GAP: structural assignability via invoke prints is thin in the existing suite.

csharp_cases! {
    func_string_to_object_covariant_return_invokes => {
        r#"System.Func<string> getString=()=>"covariant"; System.Func<object> getObject=getString; Console.WriteLine(getObject());"#,
        ["covariant"]
    };

    func_int_to_object_covariant_return_invokes => {
        r#"System.Func<int> getInt=()=>42; System.Func<object> getObject=getInt; Console.WriteLine(getObject());"#,
        ["42"]
    };

    func_bool_to_object_covariant_return_invokes => {
        r#"System.Func<bool> getBool=()=>true; System.Func<object> getObject=getBool; Console.WriteLine(getObject());"#,
        ["True"]
    };

    func_derived_to_base_covariant_return_via_subclass => {
        r#"class Animal{} class Dog:Animal{} System.Func<Dog> getDog=()=>new Dog(); System.Func<Animal> getAnimal=getDog; Console.WriteLine(getAnimal()!=null);"#,
        ["True"]
    };

    func_string_array_to_object_array_covariant => {
        r#"System.Func<string[]> getStrings=()=>new string[]{"a"}; System.Func<object[]> getObjects=getStrings; Console.WriteLine(getObjects()[0]);"#,
        ["a"]
    };

    action_object_to_string_contravariant_param_invokes => {
        r#"System.Action<object> logObject=v=>Console.WriteLine(v); System.Action<string> logString=logObject; logString("typed");"#,
        ["typed"]
    };

    action_object_to_int_contravariant_param_invokes => {
        r#"System.Action<object> logObject=v=>Console.WriteLine(v); System.Action<int> logInt=logObject; logInt(7);"#,
        ["7"]
    };

    action_base_to_derived_contravariant_via_subclass => {
        r#"class Animal{} class Dog:Animal{} System.Action<Animal> feed=a=>Console.WriteLine(a!=null); System.Action<Dog> feedDog=feed; feedDog(new Dog());"#,
        ["True"]
    };

    action_object_to_string_contravariant_accepts_null => {
        r#"System.Action<object> sink=v=>Console.WriteLine(v==null); System.Action<string> sinkString=sink; sinkString(null);"#,
        ["True"]
    };

    func_covariant_chain_two_hops_to_object => {
        r#"System.Func<string> inner=()=>"chain"; System.Func<object> mid=inner; System.Func<object> outer=mid; Console.WriteLine(outer());"#,
        ["chain"]
    };

    action_contravariant_chain_two_hops_to_string => {
        r#"System.Action<object> root=v=>Console.WriteLine(v); System.Action<object> mid=root; System.Action<string> leaf=mid; leaf("deep");"#,
        ["deep"]
    };

    func_covariant_return_preserves_string_length => {
        r#"System.Func<string> src=()=>"abcd"; System.Func<object> widened=src; Console.WriteLine(((string)widened()).Length);"#,
        ["4"]
    };

    action_contravariant_param_prints_twice => {
        r#"System.Action<object> once=v=>{Console.WriteLine(v); Console.WriteLine(v);}; System.Action<string> twice=once; twice("x");"#,
        ["x", "x"]
    };

    func_covariant_with_local_function => {
        r#"string Local()=>"local"; System.Func<string> f=Local; System.Func<object> g=f; Console.WriteLine(g());"#,
        ["local"]
    };

    action_contravariant_with_local_function => {
        r#"void Show(object o)=>Console.WriteLine(o); System.Action<object> baseAct=Show; System.Action<string> derivedAct=baseAct; derivedAct("fn");"#,
        ["fn"]
    };

    func_covariant_nullable_string_to_object => {
        r#"System.Func<string> f=()=>null; System.Func<object> g=f; Console.WriteLine(g()==null);"#,
        ["True"]
    };

    action_contravariant_int_boxed_to_object => {
        r#"System.Action<object> log=v=>Console.WriteLine(v is int); System.Action<int> logInt=log; logInt(5);"#,
        ["True"]
    };

    func_covariant_return_type_name_print => {
        r#"System.Func<string> f=()=>"name"; System.Func<object> g=f; Console.WriteLine(g().GetType().Name);"#,
        ["String"]
    };

    action_contravariant_string_is_object_check => {
        r#"System.Action<object> probe=v=>Console.WriteLine(v is string); System.Action<string> probeString=probe; probeString("ok");"#,
        ["True"]
    };

    func_covariant_delegate_stored_in_object_field => {
        r#"System.Func<string> f=()=>"field"; System.Func<object> g=f; object boxed=g; Console.WriteLine(((System.Func<object>)boxed)());"#,
        ["field"]
    };

    action_contravariant_stored_in_base_reference => {
        r#"System.Action<object> baseAct=v=>Console.WriteLine(v); System.Action<string> derivedAct=baseAct; object holder=derivedAct; ((System.Action<string>)holder)("hold");"#,
        ["hold"]
    };

    func_covariant_multiline_lambda => {
        r#"System.Func<string> narrow=()=>{return "multi";}; System.Func<object> wide=narrow; Console.WriteLine(wide());"#,
        ["multi"]
    };

    action_contravariant_multiline_lambda => {
        r#"System.Action<object> wide=v=>{Console.WriteLine(v);}; System.Action<int> narrow=wide; narrow(99);"#,
        ["99"]
    };

    func_covariant_return_empty_string => {
        r#"System.Func<string> f=()=>""; System.Func<object> g=f; Console.WriteLine(((string)g()).Length);"#,
        ["0"]
    };

    action_contravariant_empty_string_arg => {
        r#"System.Action<object> log=v=>Console.WriteLine(((string)v).Length); System.Action<string> logStr=log; logStr("");"#,
        ["0"]
    };

    func_covariant_numeric_to_object_unbox => {
        r#"System.Func<int> f=()=>123; System.Func<object> g=f; Console.WriteLine((int)g());"#,
        ["123"]
    };

    action_contravariant_object_to_string_cast => {
        r#"System.Action<object> log=v=>Console.WriteLine((string)v); System.Action<string> logStr=log; logStr("cast");"#,
        ["cast"]
    };

    func_covariant_two_independent_assignments => {
        r#"System.Func<string> a=()=>"one"; System.Func<string> b=()=>"two"; System.Func<object> ga=a; System.Func<object> gb=b; Console.WriteLine(ga()); Console.WriteLine(gb());"#,
        ["one", "two"]
    };

    action_contravariant_two_independent_assignments => {
        r#"System.Action<object> a=v=>Console.WriteLine("a"); System.Action<object> b=v=>Console.WriteLine("b"); System.Action<string> sa=a; System.Action<string> sb=b; sa(""); sb("");"#,
        ["a", "b"]
    };

    func_covariant_reassign_source => {
        r#"System.Func<object> wide=null; System.Func<string> narrow=()=>"rebind"; wide=narrow; Console.WriteLine(wide());"#,
        ["rebind"]
    };

    action_contravariant_reassign_source => {
        r#"System.Action<string> narrow=null; System.Action<object> wide=v=>Console.WriteLine(v); narrow=wide; narrow("rebind");"#,
        ["rebind"]
    };

    func_covariant_return_char_boxed => {
        r#"System.Func<char> f=()=>'Z'; System.Func<object> g=f; Console.WriteLine(g());"#,
        ["Z"]
    };

    action_contravariant_char_promoted_to_object => {
        r#"System.Action<object> log=v=>Console.WriteLine(v); System.Action<char> logChar=log; logChar('Q');"#,
        ["Q"]
    };

    func_covariant_return_double_boxed => {
        r#"System.Func<double> f=()=>3.14; System.Func<object> g=f; Console.WriteLine(g());"#,
        ["3.14"]
    };

    action_contravariant_double_to_object => {
        r#"System.Action<object> log=v=>Console.WriteLine((double)v==3.14); System.Action<double> logD=log; logD(3.14);"#,
        ["True"]
    };

    func_covariant_predicate_style_bool => {
        r#"System.Func<bool> f=()=>false; System.Func<object> g=f; Console.WriteLine(g());"#,
        ["False"]
    };

    action_contravariant_bool_boxed => {
        r#"System.Action<object> log=v=>Console.WriteLine(v); System.Action<bool> logB=log; logB(true);"#,
        ["True"]
    };

    func_covariant_return_array_length => {
        r#"System.Func<int[]> f=()=>new int[]{1,2,3}; System.Func<object> g=f; Console.WriteLine(((int[])g()).Length);"#,
        ["3"]
    };

    action_contravariant_array_as_object => {
        r#"System.Action<object> log=v=>Console.WriteLine(v is int[]); System.Action<int[]> logArr=log; logArr(new int[]{1});"#,
        ["True"]
    };

    func_covariant_method_group => {
        r#"static string Make()=>"group"; System.Func<string> narrow=Make; System.Func<object> wide=narrow; Console.WriteLine(wide());"#,
        ["group"]
    };

    action_contravariant_method_group => {
        r#"static void Print(object o)=>Console.WriteLine(o); System.Action<object> wide=Print; System.Action<string> narrow=wide; narrow("group");"#,
        ["group"]
    };

    func_covariant_invoke_twice_same_result => {
        r#"System.Func<string> f=()=>"same"; System.Func<object> g=f; Console.WriteLine(g()); Console.WriteLine(g());"#,
        ["same", "same"]
    };

    action_contravariant_invoke_twice => {
        r#"System.Action<object> w=v=>Console.WriteLine(v); System.Action<string> n=w; n("a"); n("b");"#,
        ["a", "b"]
    };

    func_covariant_return_concat => {
        r#"System.Func<string> f=()=>"ab"+"cd"; System.Func<object> g=f; Console.WriteLine(g());"#,
        ["abcd"]
    };

    action_contravariant_uppercase_via_object => {
        r#"System.Action<object> w=v=>Console.WriteLine(((string)v).ToUpper()); System.Action<string> n=w; n("hi");"#,
        ["HI"]
    };
}
