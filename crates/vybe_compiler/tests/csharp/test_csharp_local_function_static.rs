//! Local functions, `static` local functions, and outer-variable capture.
//! GAP: static-local and capture semantics need broader structural coverage.


csharp_cases! {
    local_function_basic_call_returns_square => {
        r#"int Square(int n){int Sq(int x)=>x*x; return Sq(n);} Console.WriteLine(Square(4));"#,
        ["16"]
    };

    local_function_expression_bodied => {
        r#"int Triple(int n){int T(int x)=>x*3; return T(n);} Console.WriteLine(Triple(5));"#,
        ["15"]
    };

    local_function_recursive_factorial => {
        r#"int Fact(int n){int F(int k)=>k<=1?1:k*F(k-1); return F(n);} Console.WriteLine(Fact(5));"#,
        ["120"]
    };

    local_function_recursive_fibonacci => {
        r#"int Fib(int n){int F(int k)=>k<=1?k:F(k-1)+F(k-2); return F(n);} Console.WriteLine(Fib(8));"#,
        ["21"]
    };

    local_function_captures_outer_int => {
        r#"int offset=10; int Add(int n){int B(int x)=>x+offset; return B(n);} Console.WriteLine(Add(5));"#,
        ["15"]
    };

    local_function_captures_outer_string => {
        r#"string prefix="p:"; string Tag(int n){string T(int x)=>prefix+x; return T(n);} Console.WriteLine(Tag(7));"#,
        ["p:7"]
    };

    local_function_captures_mutable_outer => {
        r#"int scale=2; int Mul(int n){int M(int x)=>x*scale; scale=3; return M(n);} Console.WriteLine(Mul(4));"#,
        ["12"]
    };

    static_local_function_no_capture_add => {
        r#"int Sum(int a,int b){static int Add(int x,int y)=>x+y; return Add(a,b);} Console.WriteLine(Sum(3,4));"#,
        ["7"]
    };

    static_local_function_no_capture_multiply => {
        r#"int Product(int a,int b){static int Mul(int x,int y)=>x*y; return Mul(a,b);} Console.WriteLine(Product(6,7));"#,
        ["42"]
    };

    static_local_function_in_static_method => {
        r#"static int Pure(int a,int b){static int Add(int x,int y)=>x+y; return Add(a,b);} Console.WriteLine(Pure(1,2));"#,
        ["3"]
    };

    static_local_function_recursive_static => {
        r#"int CountDown(int n){static int Step(int k)=>k<=0?0:1+Step(k-1); return Step(n);} Console.WriteLine(CountDown(4));"#,
        ["4"]
    };

    local_function_used_before_declaration => {
        r#"Console.WriteLine(Double(6)); int Double(int x)=>x*2;"#,
        ["12"]
    };

    local_function_nested_two_levels => {
        r#"int Outer(int n){int Mid(int x){int Inner(int y)=>y+1; return Inner(x);} return Mid(n);} Console.WriteLine(Outer(9));"#,
        ["10"]
    };

    local_function_returns_local_delegate => {
        r#"System.Func<int,int> MakeAdder(int n){int Add(int x)=>x+n; return Add;} var add5=MakeAdder(5); Console.WriteLine(add5(10));"#,
        ["15"]
    };

    local_function_capture_in_returned_delegate => {
        r#"System.Func<int,int> MakeScaler(int factor){int Scale(int x)=>x*factor; return Scale;} Console.WriteLine(MakeScaler(4)(6));"#,
        ["24"]
    };

    local_function_with_default_parameter => {
        r#"int Inc(int n){int Step(int x,int by=1)=>x+by; return Step(n,3);} Console.WriteLine(Inc(10));"#,
        ["13"]
    };

    local_function_overload_by_param_count => {
        r#"int Compute(int n){int One(int x)=>x+1; int Two(int x,int y)=>x+y; return Two(n,One(n));} Console.WriteLine(Compute(5));"#,
        ["11"]
    };

    local_function_in_loop_accumulates => {
        r#"int Sum(int n){int total=0; for(int i=1;i<=n;i++){int Add(int x)=>total+x; total=Add(i);} return total;} Console.WriteLine(Sum(3));"#,
        ["6"]
    };

    local_function_if_branch_picks_path => {
        r#"string Sign(int n){string Pos(int x)=>"+"; string Neg(int x)=>"-"; if(n>=0){return Pos(n);} return Neg(n);} Console.WriteLine(Sign(-1));"#,
        ["-"]
    };

    local_function_switch_expression => {
        r#"string Label(int n){string L(int x)=>x switch{1=>"one",2=>"two",_=>"other"}; return L(n);} Console.WriteLine(Label(2));"#,
        ["two"]
    };

    static_local_function_called_from_sibling_local => {
        r#"int Pipeline(int n){static int Double(int x)=>x*2; int Wrap(int v)=>Double(v)+1; return Wrap(n);} Console.WriteLine(Pipeline(5));"#,
        ["11"]
    };

    local_function_captures_class_field => {
        r#"class Box{public int Value=5; int Scale(int n){int S(int x)=>x*Value; return S(n);}} var b=new Box(); Console.WriteLine(b.Scale(3));"#,
        ["15"]
    };

    local_function_captures_local_struct => {
        r#"int UseStruct(){var p=new System.ValueTuple<int,int>(2,3); int Sum(int n){int S(int x)=>p.Item1+p.Item2+x; return S(n);} return Sum(1);} Console.WriteLine(UseStruct());"#,
        ["6"]
    };

    local_function_void_side_effect => {
        r#"int Run(){int acc=0; void Bump(int n){acc+=n;} Bump(2); Bump(3); return acc;} Console.WriteLine(Run());"#,
        ["5"]
    };

    static_local_function_void_no_capture => {
        r#"int Run(){int acc=0; static void Add(ref int target,int n){target+=n;} Add(ref acc,4); Add(ref acc,1); return acc;} Console.WriteLine(Run());"#,
        ["5"]
    };

    local_function_generic_style_with_object => {
        r#"string Format(int n){string F(int x)=>"n="+x; return F(n);} Console.WriteLine(Format(42));"#,
        ["n=42"]
    };

    local_function_tail_call_style_sum => {
        r#"int Sum(int n){int Loop(int i,int acc)=>i>n?acc:Loop(i+1,acc+i); return Loop(1,0);} Console.WriteLine(Sum(4));"#,
        ["10"]
    };

    local_function_capture_bool_flag => {
        r#"bool enabled=true; int Gate(int n){int G(int x)=>enabled?x:-x; return G(n);} Console.WriteLine(Gate(7));"#,
        ["7"]
    };

    local_function_capture_nullable_coalesce => {
        r#"int? maybe=8; int Coalesce(int n){int C(int x)=>x+(maybe??0); return C(n);} Console.WriteLine(Coalesce(2));"#,
        ["10"]
    };

    static_local_function_max_of_two => {
        r#"int Max(int a,int b){static int Pick(int x,int y)=>x>y?x:y; return Pick(a,b);} Console.WriteLine(Max(3,9));"#,
        ["9"]
    };

    local_function_string_builder_capture => {
        r#"string Join(int a,int b){var sb=new System.Text.StringBuilder(); string Append(int x){sb.Append(x); return sb.ToString();} Append(a); return Append(b);} Console.WriteLine(Join(1,2));"#,
        ["12"]
    };

    local_function_capture_array_length => {
        r#"int[] data={1,2,3}; int LenPlus(int n){int L(int x)=>data.Length+x; return L(n);} Console.WriteLine(LenPlus(2));"#,
        ["5"]
    };

    local_function_delegate_param => {
        r#"int Apply(int n,System.Func<int,int> op){int Wrap(int x)=>op(x)+1; return Wrap(n);} Console.WriteLine(Apply(4,x=>x*2));"#,
        ["9"]
    };

    static_local_function_identity => {
        r#"int Id(int n){static int Self(int x)=>x; return Self(n);} Console.WriteLine(Id(100));"#,
        ["100"]
    };

    local_function_capture_char => {
        r#"char ch='A'; string Show(int n){string S(int x)=>ch+""+x; return S(n);} Console.WriteLine(Show(1));"#,
        ["A1"]
    };

    local_function_nested_static_inner => {
        r#"int Calc(int n){static int Inner(int x)=>x+5; int Outer(int v)=>Inner(v)*2; return Outer(n);} Console.WriteLine(Calc(3));"#,
        ["16"]
    };

    local_function_capture_double => {
        r#"double rate=1.5; int Scale(int n){int S(int x)=>(int)(x*rate); return S(n);} Console.WriteLine(Scale(4));"#,
        ["6"]
    };

    local_function_bool_predicate => {
        r#"bool AllPositive(int a,int b){bool Check(int x,int y)=>x>0&&y>0; return Check(a,b);} Console.WriteLine(AllPositive(1,2));"#,
        ["True"]
    };

    static_local_function_min_of_two => {
        r#"int Min(int a,int b){static int Pick(int x,int y)=>x<y?x:y; return Pick(a,b);} Console.WriteLine(Min(3,9));"#,
        ["3"]
    };

    local_function_capture_enum => {
        r#"enum Mode{On,Off} Mode state=Mode.On; int Code(int n){int C(int x)=>state==Mode.On?x:0; return C(n);} Console.WriteLine(Code(5));"#,
        ["5"]
    };

    local_function_multiple_captures => {
        r#"int a=2; int b=3; int Mix(int n){int M(int x)=>a*b+x; return M(n);} Console.WriteLine(Mix(4));"#,
        ["10"]
    };

    static_local_function_abs => {
        r#"int Abs(int n){static int Pos(int x)=>x<0?-x:x; return Pos(n);} Console.WriteLine(Abs(-8));"#,
        ["8"]
    };

    local_function_capture_string_interpolation => {
        r#"string tag="id"; string Label(int n){string L(int x)=>$"{tag}:{x}"; return L(n);} Console.WriteLine(Label(9));"#,
        ["id:9"]
    };

    local_function_while_loop_with_capture => {
        r#"int Count(int n){int i=0; int acc=0; while(i<n){int Step(int x)=>acc+x; acc=Step(i+1); i++;} return acc;} Console.WriteLine(Count(3));"#,
        ["6"]
    };

    static_local_function_power => {
        r#"int Pow(int b,int e){static int Loop(int base,int exp,int acc)=>exp==0?acc:Loop(base,exp-1,acc*base); return Loop(b,e,1);} Console.WriteLine(Pow(2,4));"#,
        ["16"]
    };

    local_function_capture_list_count => {
        r#"var items=new System.Collections.Generic.List<int>{1,2,3}; int SizePlus(int n){int S(int x)=>items.Count+x; return S(n);} Console.WriteLine(SizePlus(1));"#,
        ["4"]
    };

    local_function_return_string_from_capture => {
        r#"string suffix="!"; string Exclaim(string text){string E(string t)=>t+suffix; return E(text);} Console.WriteLine(Exclaim("hi"));"#,
        ["hi!"]
    };

    static_local_function_is_even => {
        r#"bool Even(int n){static bool Check(int x)=>x%2==0; return Check(n);} Console.WriteLine(Even(6));"#,
        ["True"]
    };

    local_function_capture_byte => {
        r#"byte mask=3; int Apply(int n){int A(int x)=>x+(int)mask; return A(n);} Console.WriteLine(Apply(5));"#,
        ["8"]
    };

    local_function_two_static_siblings => {
        r#"int Combo(int n){static int A(int x)=>x+1; static int B(int x)=>x*2; return B(A(n));} Console.WriteLine(Combo(4));"#,
        ["10"]
    };

    local_function_capture_long => {
        r#"long baseVal=10000000000L; int Add(int n){int A(int x)=>x+(int)(baseVal%100); return A(n);} Console.WriteLine(Add(5));"#,
        ["5"]
    };

    static_local_function_clamp_high => {
        r#"int Clamp(int n,int max){static int Cap(int x,int m)=>x>m?m:x; return Cap(n,max);} Console.WriteLine(Clamp(15,10));"#,
        ["10"]
    };

    local_function_capture_nullable_int_has_value => {
        r#"int? opt=7; int Bump(int n){int B(int x)=>x+(opt??0); return B(n);} Console.WriteLine(Bump(1));"#,
        ["8"]
    };

    static_local_function_negate => {
        r#"int Negate(int n){static int Flip(int x)=>-x; return Flip(n);} Console.WriteLine(Negate(12));"#,
        ["-12"]
    };

    local_function_capture_params_array => {
        r#"int SumAll(int[] nums){int Total(int n){int s=0; for(int i=0;i<nums.Length;i++){s+=nums[i];} return s;} return Total(nums.Length);} Console.WriteLine(SumAll(new int[]{1,2,3}));"#,
        ["6"]
    };
}
