//! Record struct deep coverage: value equality, IEquatable, GetHashCode, with expressions, readonly/nominal forms.


csharp_cases! {
    record_struct_readonly_positional_equality => {
        r#"readonly record struct Pair(int A,int B); Console.WriteLine(new Pair(1,2)==new Pair(1,2));"#,
        ["True"]
    };

    record_struct_readonly_inequality => {
        r#"readonly record struct Pair(int A,int B); Console.WriteLine(new Pair(1,2)!=new Pair(2,1));"#,
        ["True"]
    };

    record_struct_three_field_equality => {
        r#"record struct Rgb(byte R,byte G,byte B); Console.WriteLine(new Rgb(10,20,30)==new Rgb(10,20,30));"#,
        ["True"]
    };

    record_struct_three_field_inequality_last => {
        r#"record struct Rgb(byte R,byte G,byte B); Console.WriteLine(new Rgb(10,20,30)==new Rgb(10,20,31));"#,
        ["False"]
    };

    record_struct_three_field_inequality_first => {
        r#"record struct Rgb(byte R,byte G,byte B); Console.WriteLine(new Rgb(10,20,30)==new Rgb(11,20,30));"#,
        ["False"]
    };

    record_struct_hashcode_equal => {
        r#"record struct Key(int Id); var a=new Key(7); var b=new Key(7); Console.WriteLine(a.GetHashCode()==b.GetHashCode());"#,
        ["True"]
    };

    record_struct_hashcode_differs => {
        r#"record struct Key(int Id); Console.WriteLine(new Key(1).GetHashCode()==new Key(2).GetHashCode());"#,
        ["False"]
    };

    record_struct_equals_object_equal => {
        r#"record struct Key(int Id); object o=new Key(5); Console.WriteLine(new Key(5).Equals(o));"#,
        ["True"]
    };

    record_struct_equals_object_not_equal => {
        r#"record struct Key(int Id); object o=new Key(6); Console.WriteLine(new Key(5).Equals(o));"#,
        ["False"]
    };

    record_struct_iequatable_equals => {
        r#"record struct Key(int Id); System.IEquatable<Key> e=new Key(3); Console.WriteLine(e.Equals(new Key(3)));"#,
        ["True"]
    };

    record_struct_reference_equals_false => {
        r#"record struct Key(int Id); var a=new Key(1); var b=new Key(1); Console.WriteLine(System.Object.ReferenceEquals(a,b));"#,
        ["False"]
    };

    record_struct_tostring_contains => {
        r#"record struct Tag(string Name); Console.WriteLine(new Tag("beta").ToString().Contains("beta"));"#,
        ["True"]
    };

    record_struct_single_field_equal => {
        r#"record struct Count(int N); Console.WriteLine(new Count(0)==new Count(0));"#,
        ["True"]
    };

    record_struct_negative_equal => {
        r#"record struct Offset(int Delta); Console.WriteLine(new Offset(-5)==new Offset(-5));"#,
        ["True"]
    };

    record_struct_string_empty_equal => {
        r#"record struct Label(string Text); Console.WriteLine(new Label("")==new Label(""));"#,
        ["True"]
    };

    record_struct_string_case_ineq => {
        r#"record struct Label(string Text); Console.WriteLine(new Label("A")==new Label("a"));"#,
        ["False"]
    };

    record_struct_bool_equal => {
        r#"record struct Flag(bool On); Console.WriteLine(new Flag(true)==new Flag(true));"#,
        ["True"]
    };

    record_struct_bool_ineq => {
        r#"record struct Flag(bool On); Console.WriteLine(new Flag(true)==new Flag(false));"#,
        ["False"]
    };

    record_struct_double_equal => {
        r#"record struct Rate(double V); Console.WriteLine(new Rate(2.5)==new Rate(2.5));"#,
        ["True"]
    };

    record_struct_decimal_equal => {
        r#"record struct Money(decimal A); Console.WriteLine(new Money(9.99m)==new Money(9.99m));"#,
        ["True"]
    };

    record_struct_char_equal => {
        r#"record struct Sym(char C); Console.WriteLine(new Sym('Q')==new Sym('Q'));"#,
        ["True"]
    };

    record_struct_long_equal => {
        r#"record struct Wide(long V); Console.WriteLine(new Wide(10000000000L)==new Wide(10000000000L));"#,
        ["True"]
    };

    record_struct_nullable_both_null => {
        r#"record struct Maybe(int? N); Console.WriteLine(new Maybe(null)==new Maybe(null));"#,
        ["True"]
    };

    record_struct_nullable_value_equal => {
        r#"record struct Maybe(int? N); Console.WriteLine(new Maybe(4)==new Maybe(4));"#,
        ["True"]
    };

    record_struct_nullable_null_vs_value => {
        r#"record struct Maybe(int? N); Console.WriteLine(new Maybe(null)==new Maybe(4));"#,
        ["False"]
    };

    record_struct_with_single => {
        r#"record struct Point(int X,int Y); var p=new Point(1,2); var q=p with{X=9}; Console.WriteLine(p.X); Console.WriteLine(q.X);"#,
        ["1", "9"]
    };

    record_struct_with_two => {
        r#"record struct Point(int X,int Y); var p=new Point(1,2); var q=p with{X=3,Y=4}; Console.WriteLine(p.Y); Console.WriteLine(q.X); Console.WriteLine(q.Y);"#,
        ["2", "3", "4"]
    };

    record_struct_with_chain => {
        r#"record struct Box(int W,int H); var a=new Box(1,1); var b=a with{W=2}; var c=b with{H=3}; Console.WriteLine(a.W); Console.WriteLine(c.W); Console.WriteLine(c.H);"#,
        ["1", "2", "3"]
    };

    record_struct_copy_with => {
        r#"record struct Count(int N); var a=new Count(5); var b=a; b=b with{N=99}; Console.WriteLine(a.N); Console.WriteLine(b.N);"#,
        ["5", "99"]
    };

    record_struct_with_readonly => {
        r#"readonly record struct Size(int W,int H); var s=new Size(2,3); var t=s with{H=8}; Console.WriteLine(s.H); Console.WriteLine(t.H);"#,
        ["3", "8"]
    };

    record_struct_with_nominal => {
        r#"record struct Config{public int Port{get;init;}=80;} var c=new Config{Port=8080}; var d=c with{Port=443}; Console.WriteLine(c.Port); Console.WriteLine(d.Port);"#,
        ["8080", "443"]
    };

    record_struct_with_preserves_init => {
        r#"record struct Pair{public int A{get;init;} public int B{get;init;}} var p=new Pair{A=1,B=2}; var q=p with{A=9}; Console.WriteLine(q.B);"#,
        ["2"]
    };

    record_struct_positional_init_with => {
        r#"record struct User(string Name){public int Age{get;init;}} var u=new User("Ada"){Age=30}; var v=u with{Age=31}; Console.WriteLine(u.Age); Console.WriteLine(v.Age);"#,
        ["30", "31"]
    };

    record_struct_is_value_type => {
        r#"record struct Coord(int X,int Y); Console.WriteLine(typeof(Coord).IsValueType);"#,
        ["True"]
    };

    record_struct_pass_method => {
        r#"record struct V(int N); int Read(V v)=>v.N; Console.WriteLine(Read(new V(12)));"#,
        ["12"]
    };

    record_struct_return_method => {
        r#"record struct V(int N); V Make()=>new V(7); Console.WriteLine(Make().N);"#,
        ["7"]
    };

    record_struct_array_index => {
        r#"record struct V(int N); var arr=new[]{new V(1),new V(2)}; Console.WriteLine(arr[1].N);"#,
        ["2"]
    };

    record_struct_foreach_sum => {
        r#"record struct V(int N); var sum=0; foreach(var v in new[]{new V(1),new V(2),new V(3)}) sum+=v.N; Console.WriteLine(sum);"#,
        ["6"]
    };

    record_struct_deconstruct => {
        r#"record struct Vec(int X,int Y); var (x,y)=new Vec(3,4); Console.WriteLine(x+y);"#,
        ["7"]
    };

    record_struct_computed_property => {
        r#"record struct Rect(int W,int H){public int Area=>W*H;} Console.WriteLine(new Rect(3,4).Area);"#,
        ["12"]
    };

    record_struct_custom_method => {
        r#"record struct V(int N){public int Twice()=>N*2;} Console.WriteLine(new V(6).Twice());"#,
        ["12"]
    };

    record_struct_static_factory => {
        r#"record struct V(int N){public static V Zero()=>new V(0);} Console.WriteLine(V.Zero().N);"#,
        ["0"]
    };

    record_struct_enum_equal => {
        r#"enum Level{Low,High} record struct Job(Level Tier); Console.WriteLine(new Job(Level.High)==new Job(Level.High));"#,
        ["True"]
    };

    record_struct_enum_with => {
        r#"enum Level{Low,High} record struct Job(Level Tier); var j=new Job(Level.Low); var k=j with{Tier=Level.High}; Console.WriteLine(k.Tier);"#,
        ["High"]
    };

    record_struct_max_int_equal => {
        r#"record struct Edge(int V); Console.WriteLine(new Edge(int.MaxValue)==new Edge(int.MaxValue));"#,
        ["True"]
    };

    record_struct_min_int_equal => {
        r#"record struct Edge(int V); Console.WriteLine(new Edge(int.MinValue)==new Edge(int.MinValue));"#,
        ["True"]
    };

    record_struct_independent_instances => {
        r#"record struct V(int N); var a=new V(1); var b=new V(2); var c=a with{N=5}; Console.WriteLine(b.N); Console.WriteLine(c.N);"#,
        ["2", "5"]
    };

    record_struct_equal_after_with_same => {
        r#"record struct V(int N); var a=new V(1); var b=a with{N=1}; Console.WriteLine(a==b);"#,
        ["True"]
    };

    record_struct_not_equal_after_with => {
        r#"record struct V(int N); var a=new V(1); var b=a with{N=2}; Console.WriteLine(a==b);"#,
        ["False"]
    };

    record_struct_default_init => {
        r#"record struct Tag{public string Name{get;init;}="none";} Console.WriteLine(new Tag().Name);"#,
        ["none"]
    };

    record_struct_float_equal => {
        r#"record struct Sample(float R); Console.WriteLine(new Sample(1.5f)==new Sample(1.5f));"#,
        ["True"]
    };

    record_struct_byte_ineq => {
        r#"record struct ByteVal(byte B); Console.WriteLine(new ByteVal(1)==new ByteVal(2));"#,
        ["False"]
    };

    record_struct_hash_after_copy => {
        r#"record struct Key(int Id); var a=new Key(9); var b=a; Console.WriteLine(a.GetHashCode()==b.GetHashCode());"#,
        ["True"]
    };

    record_struct_equals_null => {
        r#"record struct Key(int Id); Console.WriteLine(new Key(1).Equals(null));"#,
        ["False"]
    };

    record_struct_custom_tostring => {
        r#"record struct Tag(string Name){public override string ToString()=>"Tag:"+Name;} Console.WriteLine(new Tag("x"));"#,
        ["Tag:x"]
    };
}
