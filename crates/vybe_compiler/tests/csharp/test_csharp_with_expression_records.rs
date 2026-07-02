//! With expressions on records: positional/nominal, nested with, derived records, and reference semantics.


csharp_cases! {
    with_positional_single_field => {
        r#"record Point(int X,int Y); var q=(new Point(1,2)) with{X=10}; Console.WriteLine(q.X); Console.WriteLine(q.Y);"#,
        ["10", "2"]
    };

    with_positional_two_fields => {
        r#"record Point(int X,int Y); var q=(new Point(1,2)) with{X=3,Y=4}; Console.WriteLine(q.X); Console.WriteLine(q.Y);"#,
        ["3", "4"]
    };

    with_positional_original_unchanged => {
        r#"record Point(int X,int Y); var p=new Point(1,2); var q=p with{X=9}; Console.WriteLine(p.X);"#,
        ["1"]
    };

    with_chained_three_steps => {
        r#"record Box(int W,int H,int D); var a=new Box(1,2,3); var b=a with{W=4}; var c=b with{H=5}; var d=c with{D=6}; Console.WriteLine(a.W); Console.WriteLine(d.W); Console.WriteLine(d.H); Console.WriteLine(d.D);"#,
        ["1", "4", "5", "6"]
    };

    with_nominal_init_port => {
        r#"record Config{public string Host{get;init;}="localhost"; public int Port{get;init;}=80;} var p=(new Config()) with{Port=443}; Console.WriteLine(p.Port);"#,
        ["443"]
    };

    with_nominal_host_only => {
        r#"record Config{public string Host{get;init;} public int Port{get;init;}} var c=new Config{Host="a",Port=1}; var d=c with{Host="b"}; Console.WriteLine(c.Host); Console.WriteLine(d.Host);"#,
        ["a", "b"]
    };

    with_nominal_two_inits => {
        r#"record Theme{public string Name{get;init;} public int Ver{get;init;}} var u=(new Theme{Name="dark",Ver=1}) with{Name="light",Ver=2}; Console.WriteLine(u.Name); Console.WriteLine(u.Ver);"#,
        ["light", "2"]
    };

    with_positional_plus_init => {
        r#"record User(string Name){public int Age{get;init;}} var v=(new User("Ada"){Age=20}) with{Age=21}; Console.WriteLine(v.Age);"#,
        ["21"]
    };

    with_string_field => {
        r#"record Tag(string Name); var n=(new Tag("old")) with{Name="new"}; Console.WriteLine(n.Name);"#,
        ["new"]
    };

    with_bool_toggle => {
        r#"record Flag(bool On); var g=(new Flag(false)) with{On=true}; Console.WriteLine(g.On);"#,
        ["True"]
    };

    with_decimal_field => {
        r#"record Price(decimal A); var q=(new Price(1.5m)) with{A=9.99m}; Console.WriteLine(q.A);"#,
        ["9.99"]
    };

    with_double_field => {
        r#"record Rate(double V); var s=(new Rate(1.1)) with{V=2.2}; Console.WriteLine(s.V);"#,
        ["2.2"]
    };

    with_char_field => {
        r#"record Sym(char C); var b=(new Sym('a')) with{C='z'}; Console.WriteLine(b.C);"#,
        ["z"]
    };

    with_enum_field => {
        r#"enum Mode{Off,On} record State(Mode M); var t=(new State(Mode.Off)) with{M=Mode.On}; Console.WriteLine(t.M);"#,
        ["On"]
    };

    with_nullable_to_null => {
        r#"record Maybe(int? N); var z=(new Maybe(5)) with{N=null}; Console.WriteLine(z.N.HasValue);"#,
        ["False"]
    };

    with_nullable_to_value => {
        r#"record Maybe(int? N); var v=(new Maybe(null)) with{N=7}; Console.WriteLine(v.N);"#,
        ["7"]
    };

    with_three_positional_one => {
        r#"record Triple(int A,int B,int C); var u=(new Triple(1,2,3)) with{B=9}; Console.WriteLine(u.B);"#,
        ["9"]
    };

    with_three_positional_all => {
        r#"record Triple(int A,int B,int C); var u=(new Triple(1,2,3)) with{A=4,B=5,C=6}; Console.WriteLine(u.A+u.B+u.C);"#,
        ["15"]
    };

    with_four_positional => {
        r#"record Quad(int A,int B,int C,int D); var r=(new Quad(1,2,3,4)) with{D=10}; Console.WriteLine(r.D);"#,
        ["10"]
    };

    with_class_record_not_same_ref => {
        r#"record Node(int Id); var a=new Node(1); var b=a with{Id=2}; Console.WriteLine(System.Object.ReferenceEquals(a,b));"#,
        ["False"]
    };

    with_class_record_equal_same_values => {
        r#"record Node(int Id); var a=new Node(1); var b=a with{Id=1}; Console.WriteLine(a==b);"#,
        ["True"]
    };

    with_class_record_not_equal => {
        r#"record Node(int Id); var a=new Node(1); var b=a with{Id=2}; Console.WriteLine(a==b);"#,
        ["False"]
    };

    with_two_branches_same_source => {
        r#"record Pair(int A,int B); var p=new Pair(1,2); var x=p with{A=9}; var y=p with{B=8}; Console.WriteLine(x.A); Console.WriteLine(y.B);"#,
        ["9", "8"]
    };

    with_derived_derived_field => {
        r#"record Animal(string Name); record Dog(string Name,string Breed):Animal(Name); var k=(new Dog("Rex","Lab")) with{Breed="Pug"}; Console.WriteLine(k.Breed);"#,
        ["Pug"]
    };

    with_derived_base_field => {
        r#"record Animal(string Name); record Dog(string Name,string Breed):Animal(Name); var k=(new Dog("Rex","Lab")) with{Name="Max"}; Console.WriteLine(k.Name);"#,
        ["Max"]
    };

    with_nested_outer_name => {
        r#"record Address(string City); record Person(string Name,Address Home); var q=(new Person("Ann",new Address("Oslo"))) with{Name="Bob"}; Console.WriteLine(q.Name);"#,
        ["Bob"]
    };

    with_nested_replace_inner => {
        r#"record Address(string City); record Person(string Name,Address Home); var p=new Person("Ann",new Address("Oslo")); var q=p with{Home=new Address("Paris")}; Console.WriteLine(q.Home.City);"#,
        ["Paris"]
    };

    with_nested_inner_with => {
        r#"record Address(string City); record Person(string Name,Address Home); var p=new Person("Ann",new Address("Oslo")); var q=p with{Home=p.Home with{City="Paris"}}; Console.WriteLine(q.Home.City);"#,
        ["Paris"]
    };

    with_nested_triple_inner => {
        r#"record Zip(string Code); record Address(string City,Zip Z); record Person(string Name,Address Home); var p=new Person("A",new Address("Oslo",new Zip("01"))); var q=p with{Home=p.Home with{Z=p.Home.Z with{Code="02"}}}; Console.WriteLine(q.Home.Z.Code);"#,
        ["02"]
    };

    with_record_method_after => {
        r#"record Counter(int N){public int Next()=>N+1;} var d=(new Counter(1)) with{N=5}; Console.WriteLine(d.Next());"#,
        ["6"]
    };

    with_zero_int => {
        r#"record V(int N); var z=(new V(5)) with{N=0}; Console.WriteLine(z.N);"#,
        ["0"]
    };

    with_negative_int => {
        r#"record V(int N); var n=(new V(5)) with{N=-1}; Console.WriteLine(n.N);"#,
        ["-1"]
    };

    with_max_int => {
        r#"record V(int N); var m=(new V(1)) with{N=int.MaxValue}; Console.WriteLine(m.N==int.MaxValue);"#,
        ["True"]
    };

    with_string_empty => {
        r#"record Label(string T); var e=(new Label("a")) with{T=""}; Console.WriteLine(e.T.Length);"#,
        ["0"]
    };

    with_preserves_other_nominal => {
        r#"record Pair{public int A{get;init;} public int B{get;init;}} var q=(new Pair{A=1,B=2}) with{A=9}; Console.WriteLine(q.B);"#,
        ["2"]
    };

    with_inline_in_expression => {
        r#"record V(int N); Console.WriteLine((new V(2) with{N=7}).N);"#,
        ["7"]
    };

    with_long_field => {
        r#"record Wide(long V); var x=(new Wide(10L)) with{V=20L}; Console.WriteLine(x.V);"#,
        ["20"]
    };

    with_byte_field => {
        r#"record ByteBox(byte B); var c=(new ByteBox(1)) with{B=255}; Console.WriteLine(c.B);"#,
        ["255"]
    };

    with_short_field => {
        r#"record ShortBox(short S); var t=(new ShortBox(1)) with{S=1000}; Console.WriteLine(t.S);"#,
        ["1000"]
    };

    with_float_field => {
        r#"record Sample(float R); var t=(new Sample(1.0f)) with{R=2.5f}; Console.WriteLine(t.R);"#,
        ["2.5"]
    };

    with_list_element => {
        r#"record V(int N); var list=new System.Collections.Generic.List<V>{new V(1),new V(2)}; list[1]=list[1] with{N=9}; Console.WriteLine(list[1].N);"#,
        ["9"]
    };

    with_mutable_separate_instance => {
        r#"record Box{public int V{get;set;}} var b=(new Box{V=1}) with{V=2}; Console.WriteLine(b.V);"#,
        ["2"]
    };

    with_after_mutate_original => {
        r#"record Box{public int V{get;set;}} var a=new Box{V=1}; a.V=3; var b=a with{V=4}; Console.WriteLine(b.V);"#,
        ["4"]
    };

    with_tostring_after => {
        r#"record Tag(string Name); var u=(new Tag("a")) with{Name="b"}; Console.WriteLine(u.ToString().Contains("b"));"#,
        ["True"]
    };

    with_derived_tostring => {
        r#"record Animal(string Name); record Cat(string Name,string Color):Animal(Name); var d=(new Cat("M","W")) with{Color="B"}; Console.WriteLine(d.ToString().Contains("B"));"#,
        ["True"]
    };

    with_double_nested_independent => {
        r#"record Pair(int A,int B); var p=new Pair(1,1); var a=p with{A=2}; var b=p with{B=3}; Console.WriteLine(a.A); Console.WriteLine(b.B);"#,
        ["2", "3"]
    };

    with_nominal_chain => {
        r#"record C{public int A{get;init;} public int B{get;init;}} var e=((new C{A=1,B=2}) with{A=3}) with{B=4}; Console.WriteLine(e.A); Console.WriteLine(e.B);"#,
        ["3", "4"]
    };

    with_hash_changes => {
        r#"record Key(int Id); var a=new Key(1); var b=a with{Id=2}; Console.WriteLine(a.GetHashCode()==b.GetHashCode());"#,
        ["False"]
    };

    with_hash_same_when_equal => {
        r#"record Key(int Id); var a=new Key(1); var b=a with{Id=1}; Console.WriteLine(a.GetHashCode()==b.GetHashCode());"#,
        ["True"]
    };

    with_positional_preserves_string => {
        r#"record Pair(string S,int N); var q=(new Pair("x",1)) with{N=2}; Console.WriteLine(q.S);"#,
        ["x"]
    };
}
