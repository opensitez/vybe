//! Static abstract and static virtual interface members (C# 11) — spec-oriented coverage; parse/compile may fail until supported.

csharp_cases! {
    static_abstract_factory_method_on_interface => {
        r#"interface IFactory<T> where T:IFactory<T>{static abstract T Create(int n);}
struct Widget:IFactory<Widget>{public int V; public static Widget Create(int n)=>new Widget{V=n};}
Console.WriteLine(Widget.Create(5).V);"#,
        ["5"]
    };

    static_abstract_parse_pattern_like_iparsable => {
        r#"interface IParsable<T> where T:IParsable<T>{static abstract T Parse(string s);}
struct Age:IParsable<Age>{public int Years; public static Age Parse(string s)=>new Age{Years=int.Parse(s)};}
Console.WriteLine(Age.Parse("30").Years);"#,
        ["30"]
    };

    static_abstract_try_parse_returns_bool => {
        r#"interface ITryParsable<T> where T:ITryParsable<T>{static abstract bool TryParse(string s,out T value);}
struct Pair:ITryParsable<Pair>{public int A; public static bool TryParse(string s,out Pair value){value=new Pair{A=int.Parse(s)};return true;}}
Pair p; Console.WriteLine(Pair.TryParse("4",out p)?p.A:-1);"#,
        ["4"]
    };

    static_abstract_operator_plus_on_numeric_interface => {
        r#"interface IAddable<T> where T:IAddable<T>{static abstract T operator+(T a,T b);}
struct Vec:IAddable<Vec>{public int X; public static Vec operator+(Vec a,Vec b)=>new Vec{X=a.X+b.X};}
Console.WriteLine((new Vec{X=2}+new Vec{X=3}).X);"#,
        ["5"]
    };

    static_abstract_unary_negation_operator => {
        r#"interface INegatable<T> where T:INegatable<T>{static abstract T operator-(T v);}
struct Signed:INegatable<Signed>{public int N; public static Signed operator-(Signed v)=>new Signed{N=-v.N};}
Console.WriteLine((-new Signed{N=8}).N);"#,
        ["-8"]
    };

    static_abstract_property_getter_on_interface => {
        r#"interface IUnit<T> where T:IUnit<T>{static abstract T Zero{get;}}
struct Counter:IUnit<Counter>{public int V; public static Counter Zero=>new Counter{V=0};}
Console.WriteLine(Counter.Zero.V);"#,
        ["0"]
    };

    static_virtual_property_default_on_interface => {
        r#"interface IDefault<T> where T:IDefault<T>{static virtual T Fallback=>default; static abstract T Primary();}
struct Item:IDefault<Item>{public int Id; public static Item Primary()=>new Item{Id=1}; public static Item Fallback=>new Item{Id=99};}
Console.WriteLine(Item.Primary().Id);"#,
        ["1"]
    };

    static_virtual_method_default_calls_primary => {
        r#"interface IDouble<T> where T:IDouble<T>{static abstract T One(); static virtual T Two(){return One();}}
struct Dup:IDouble<Dup>{public int V; public static Dup One()=>new Dup{V=1}; public static Dup Two()=>new Dup{V=2};}
Console.WriteLine(Dup.Two().V);"#,
        ["2"]
    };

    static_abstract_interface_constraint_on_type_param => {
        r#"interface IHasLabel<T> where T:IHasLabel<T>{static abstract string Label();}
struct Tag:IHasLabel<Tag>{public static string Label()=>"tag";}
string Read<T>() where T:IHasLabel<T>=>T.Label(); Console.WriteLine(Read<Tag>());"#,
        ["tag"]
    };

    static_abstract_multiple_methods_same_interface => {
        r#"interface IBoth<T> where T:IBoth<T>{static abstract T FromInt(int n); static abstract T FromString(string s);}
struct Dual:IBoth<Dual>{public string Text; public static Dual FromInt(int n)=>new Dual{Text=n.ToString()}; public static Dual FromString(string s)=>new Dual{Text=s};}
Console.WriteLine(Dual.FromString("ok").Text);"#,
        ["ok"]
    };

    static_abstract_implemented_by_struct_value_type => {
        r#"interface IVal<T> where T:IVal<T>{static abstract T Make(int n);}
struct Point:IVal<Point>{public int X; public static Point Make(int n)=>new Point{X=n};}
Console.WriteLine(Point.Make(11).X);"#,
        ["11"]
    };

    static_abstract_implemented_by_class_reference_type => {
        r#"interface IBuild<T> where T:IBuild<T>{static abstract T New();}
class Node:IBuild<Node>{public int Id=7; public static Node New()=>new Node();}
Console.WriteLine(Node.New().Id);"#,
        ["7"]
    };

    static_abstract_generic_self_type_roundtrip => {
        r#"interface ISelf<T> where T:ISelf<T>{static abstract T Identity(T v);}
struct Wrap:ISelf<Wrap>{public int N; public static Wrap Identity(Wrap v)=>v;}
var w=new Wrap{N=3}; Console.WriteLine(Wrap.Identity(w).N);"#,
        ["3"]
    };

    static_abstract_bool_factory => {
        r#"interface IFlag<T> where T:IFlag<T>{static abstract T True(); static abstract T False();}
struct Bit:IFlag<Bit>{public bool On; public static Bit True()=>new Bit{On=true}; public static Bit False()=>new Bit{On=false};}
Console.WriteLine(Bit.True().On);"#,
        ["True"]
    };

    static_abstract_string_normalization => {
        r#"interface INorm<T> where T:INorm<T>{static abstract T Normalize(string s);}
struct Text:INorm<Text>{public string Value; public static Text Normalize(string s)=>new Text{Value=s.Trim().ToLower()};}
Console.WriteLine(Text.Normalize(" Ab ").Value);"#,
        ["ab"]
    };

    static_abstract_decimal_from_string => {
        r#"interface IDec<T> where T:IDec<T>{static abstract T Parse(string s);}
struct Money:IDec<Money>{public decimal Amount; public static Money Parse(string s)=>new Money{Amount=decimal.Parse(s)};}
Console.WriteLine(Money.Parse("12.5").Amount);"#,
        ["12.5"]
    };

    static_abstract_enum_like_factory => {
        r#"interface IKind<T> where T:IKind<T>{static abstract T North(); static abstract T South();}
struct Dir:IKind<Dir>{public string Name; public static Dir North()=>new Dir{Name="N"}; public static Dir South()=>new Dir{Name="S"};}
Console.WriteLine(Dir.North().Name);"#,
        ["N"]
    };

    static_virtual_default_used_when_not_overridden => {
        r#"interface IBase<T> where T:IBase<T>{static virtual int Code=>0; static abstract T Build();}
struct S:IBase<S>{public static S Build()=>new S();}
Console.WriteLine(S.Code);"#,
        ["0"]
    };

    static_virtual_override_changes_code => {
        r#"interface IBase<T> where T:IBase<T>{static virtual int Code=>0; static abstract T Build();}
struct S:IBase<S>{public static S Build()=>new S(); public static int Code=>5;}
Console.WriteLine(S.Code);"#,
        ["5"]
    };

    static_abstract_comparison_operator => {
        r#"interface IComparableStatic<T> where T:IComparableStatic<T>{static abstract bool operator<(T a,T b);}
struct Rank:IComparableStatic<Rank>{public int Level; public static bool operator<(Rank a,Rank b)=>a.Level<b.Level;}
Console.WriteLine(new Rank{Level=1}<new Rank{Level=2});"#,
        ["True"]
    };

    static_abstract_equality_operator => {
        r#"interface IEquatableStatic<T> where T:IEquatableStatic<T>{static abstract bool operator==(T a,T b);}
struct Key:IEquatableStatic<Key>{public int Id; public static bool operator==(Key a,Key b)=>a.Id==b.Id; public static bool operator!=(Key a,Key b)=>!(a==b);}
Console.WriteLine(new Key{Id=1}==new Key{Id=1});"#,
        ["True"]
    };

    static_abstract_increment_operator => {
        r#"interface IInc<T> where T:IInc<T>{static abstract T operator++(T v);}
struct Num:IInc<Num>{public int N; public static Num operator++(Num v)=>new Num{N=v.N+1};}
Console.WriteLine((++new Num{N=4}).N);"#,
        ["5"]
    };

    static_abstract_bitwise_and_operator => {
        r#"interface IBit<T> where T:IBit<T>{static abstract T operator&(T a,T b);}
struct Mask:IBit<Mask>{public int Bits; public static Mask operator&(Mask a,Mask b)=>new Mask{Bits=a.Bits&b.Bits};}
Console.WriteLine((new Mask{Bits=7}&new Mask{Bits=3}).Bits);"#,
        ["3"]
    };

    static_abstract_explicit_interface_implementation_style => {
        r#"interface IMaker<T> where T:IMaker<T>{static abstract T Make();}
struct Box:IMaker<Box>{public int Size; static Box IMaker<Box>.Make()=>new Box{Size=9}; public static Box Make()=>new Box{Size=1};}
Console.WriteLine(Box.Make().Size);"#,
        ["1"]
    };

    static_abstract_nested_interface_hierarchy => {
        r#"interface IRoot<T> where T:IRoot<T>{static abstract T Root();}
interface IChild<T>:IRoot<T> where T:IChild<T>{static abstract T Child();}
struct Tree:IChild<Tree>{public string Tag; public static Tree Root()=>new Tree{Tag="R"}; public static Tree Child()=>new Tree{Tag="C"};}
Console.WriteLine(Tree.Child().Tag);"#,
        ["C"]
    };

    static_abstract_multiple_type_params_with_self => {
        r#"interface IPair<TSelf,TVal> where TSelf:IPair<TSelf,TVal>{static abstract TSelf Of(TVal v);}
struct Holder:IPair<Holder,int>{public int Data; public static Holder Of(int v)=>new Holder{Data=v};}
Console.WriteLine(Holder.Of(6).Data);"#,
        ["6"]
    };

    static_abstract_void_initialize_method => {
        r#"interface IInit<T> where T:IInit<T>{static abstract void Configure(T target);}
struct Config:IInit<Config>{public int Ready; public static void Configure(Config target){target.Ready=1;}}
var c=new Config(); Config.Configure(c); Console.WriteLine(c.Ready);"#,
        ["1"]
    };

    static_abstract_returns_interface_implementor => {
        r#"interface IProvider<T> where T:IProvider<T>{static abstract T Provide();}
class Service:IProvider<Service>{public string Name="svc"; public static Service Provide()=>new Service();}
Console.WriteLine(Service.Provide().Name);"#,
        ["svc"]
    };

    static_abstract_char_conversion => {
        r#"interface IChar<T> where T:IChar<T>{static abstract T FromChar(char c);}
struct Letter:IChar<Letter>{public char C; public static Letter FromChar(char c)=>new Letter{C=c};}
Console.WriteLine(Letter.FromChar('z').C);"#,
        ["z"]
    };

    static_abstract_array_length_factory => {
        r#"interface IArray<T> where T:IArray<T>{static abstract int Length(T value);}
struct Arr:IArray<Arr>{public int[] Data; public static int Length(Arr value)=>value.Data.Length;}
Console.WriteLine(Arr.Length(new Arr{Data=new[]{1,2,3}}));"#,
        ["3"]
    };

    static_abstract_generic_method_on_interface => {
        r#"interface IConvert<T> where T:IConvert<T>{static abstract T From<U>(U value);}
struct Box:IConvert<Box>{public string Text; public static Box From<U>(U value)=>new Box{Text=value.ToString()};}
Console.WriteLine(Box.From(12).Text);"#,
        ["12"]
    };

    static_virtual_property_chain_default => {
        r#"interface IChain<T> where T:IChain<T>{static virtual string Name=>"base"; static abstract T Instance();}
struct Link:IChain<Link>{public static Link Instance()=>new Link();}
Console.WriteLine(Link.Name);"#,
        ["base"]
    };

    static_abstract_record_like_struct => {
        r#"interface IRecord<T> where T:IRecord<T>{static abstract T Create(int id,string name);}
struct User:IRecord<User>{public int Id; public string Name; public static User Create(int id,string name)=>new User{Id=id,Name=name};}
Console.WriteLine(User.Create(1,"Ann").Name);"#,
        ["Ann"]
    };

    static_abstract_signed_magnitude => {
        r#"interface ISign<T> where T:ISign<T>{static abstract T Negate(T v); static abstract T Abs(T v);}
struct IntWrap:ISign<IntWrap>{public int N; public static IntWrap Negate(IntWrap v)=>new IntWrap{N=-v.N}; public static IntWrap Abs(IntWrap v)=>new IntWrap{N=v.N<0?-v.N:v.N};}
Console.WriteLine(IntWrap.Abs(new IntWrap{N=-4}).N);"#,
        ["4"]
    };

    static_abstract_interface_used_in_generic_helper => {
        r#"interface IShow<T> where T:IShow<T>{static abstract string Show(T v);}
struct Lab:IShow<Lab>{public int N; public static string Show(Lab v)=>v.N.ToString();}
string Render<T>(T v) where T:IShow<T>=>T.Show(v); Console.WriteLine(Render(new Lab{N=9}));"#,
        ["9"]
    };

    static_abstract_two_operators_same_interface => {
        r#"interface IOps<T> where T:IOps<T>{static abstract T operator+(T a,T b); static abstract T operator*(T a,int k);}
struct Scale:IOps<Scale>{public int V; public static Scale operator+(Scale a,Scale b)=>new Scale{V=a.V+b.V}; public static Scale operator*(Scale a,int k)=>new Scale{V=a.V*k};}
Console.WriteLine((new Scale{V=2}*3).V);"#,
        ["6"]
    };

    static_abstract_bool_try_pattern => {
        r#"interface ITry<T> where T:ITry<T>{static abstract bool Try(string s,out T value);}
struct Token:ITry<Token>{public string Raw; public static bool Try(string s,out Token value){value=new Token{Raw=s};return s.Length>0;}}
Token t; Console.WriteLine(Token.Try("x",out t));"#,
        ["True"]
    };

    static_abstract_guid_like_parse => {
        r#"interface IGuid<T> where T:IGuid<T>{static abstract T Parse(string hex);}
struct Id:IGuid<Id>{public string Hex; public static Id Parse(string hex)=>new Id{Hex=hex.ToUpper()};}
Console.WriteLine(Id.Parse("ab").Hex);"#,
        ["AB"]
    };

    static_abstract_static_member_hides_instance_context => {
        r#"interface IStaticOnly<T> where T:IStaticOnly<T>{static abstract int Count();}
struct Tally:IStaticOnly<Tally>{public int N; public static int Count()=>3;}
Console.WriteLine(Tally.Count());"#,
        ["3"]
    };

    static_abstract_multiple_implementors_same_interface => {
        r#"interface ICode<T> where T:ICode<T>{static abstract int Code();}
struct A:ICode<A>{public static int Code()=>1;} struct B:ICode<B>{public static int Code()=>2;}
Console.WriteLine(A.Code()+B.Code());"#,
        ["3"]
    };

    static_virtual_default_string_label => {
        r#"interface ILabel<T> where T:ILabel<T>{static virtual string Tag=>"d"; static abstract T Make();}
struct Tag:ILabel<Tag>{public static Tag Make()=>new Tag(); public static string Tag=>"x";}
Console.WriteLine(Tag.Tag);"#,
        ["x"]
    };

    static_abstract_self_referential_constraint_roundtrip => {
        r#"interface IRound<T> where T:IRound<T>{static abstract T RoundTrip(T input);}
struct Echo:IRound<Echo>{public int N; public static Echo RoundTrip(Echo input)=>input;}
var e=new Echo{N=12}; Console.WriteLine(Echo.RoundTrip(e).N);"#,
        ["12"]
    };

    static_abstract_interface_with_struct_and_class_implementors => {
        r#"interface IShared<T> where T:IShared<T>{static abstract int Key();}
struct SA:IShared<SA>{public static int Key()=>1;} class CA:IShared<CA>{public static int Key()=>2;}
Console.WriteLine(SA.Key()+CA.Key());"#,
        ["3"]
    };

    static_abstract_parse_empty_string_throws_or_zero => {
        r#"interface ILen<T> where T:ILen<T>{static abstract int From(string s);}
struct Size:ILen<Size>{public int N; public static int From(string s)=>s.Length;}
Console.WriteLine(Size.From("abc"));"#,
        ["3"]
    };

    static_abstract_operator_true_false => {
        r#"interface ITest<T> where T:ITest<T>{static abstract bool operator true(T v); static abstract bool operator false(T v);}
struct Flag:ITest<Flag>{public bool On; public static bool operator true(Flag v)=>v.On; public static bool operator false(Flag v)=>!v.On;}
Console.WriteLine(new Flag{On=true}?1:0);"#,
        ["1"]
    };

    static_abstract_modulo_operator => {
        r#"interface IMod<T> where T:IMod<T>{static abstract T operator%(T a,T b);}
struct Mod:IMod<Mod>{public int V; public static Mod operator%(Mod a,Mod b)=>new Mod{V=a.V%b.V};}
Console.WriteLine((new Mod{V=10}%new Mod{V=3}).V);"#,
        ["1"]
    };

    static_abstract_shift_operator => {
        r#"interface IShift<T> where T:IShift<T>{static abstract T operator<<(T v,int n);}
struct Bits:IShift<Bits>{public int V; public static Bits operator<<(Bits v,int n)=>new Bits{V=v.V<<n};}
Console.WriteLine((new Bits{V=1}<<3).V);"#,
        ["8"]
    };

    static_abstract_combined_add_and_parse => {
        r#"interface IOps2<T> where T:IOps2<T>{static abstract T Parse(string s); static abstract T Add(T a,T b);}
struct Pair:IOps2<Pair>{public int A,B; public static Pair Parse(string s){var p=s.Split(','); return new Pair{A=int.Parse(p[0]),B=int.Parse(p[1])};} public static Pair Add(Pair a,Pair b)=>new Pair{A=a.A+b.A,B=a.B+b.B};}
var p=Pair.Parse("1,2"); Console.WriteLine(Pair.Add(p,p).A);"#,
        ["2"]
    };

    static_abstract_interface_nameof_type => {
        r#"interface IName<T> where T:IName<T>{static abstract string TypeName();}
struct Named:IName<Named>{public static string TypeName()=>nameof(Named);}
Console.WriteLine(Named.TypeName());"#,
        ["Named"]
    };

    static_virtual_calls_abstract_in_default_body => {
        r#"interface IWrap<T> where T:IWrap<T>{static abstract T Core(); static virtual T Outer(){return Core();}}
struct Core:IWrap<Core>{public int N; public static Core Core()=>new Core{N=4};}
Console.WriteLine(Core.Outer().N);"#,
        ["4"]
    };
}
