//! Deconstruction: tuples, positional records, custom Deconstruct methods, var (x,y)= syntax.


csharp_cases! {
    tuple_deconstruct_two => {
        r#"var (a,b)=(1,2); Console.WriteLine(a+b);"#,
        ["3"]
    };

    tuple_deconstruct_three => {
        r#"var (a,b,c)=(1,2,3); Console.WriteLine(a+b+c);"#,
        ["6"]
    };

    tuple_deconstruct_four => {
        r#"var (a,b,c,d)=(1,2,3,4); Console.WriteLine(a+b+c+d);"#,
        ["10"]
    };

    tuple_var_syntax => {
        r#"var (x,y)=(5,7); Console.WriteLine(x); Console.WriteLine(y);"#,
        ["5", "7"]
    };

    tuple_assign_existing => {
        r#"int x=0,y=0; (x,y)=(9,1); Console.WriteLine(x); Console.WriteLine(y);"#,
        ["9", "1"]
    };

    tuple_string_int => {
        r#"var (name,n)=("Ada",42); Console.WriteLine(name); Console.WriteLine(n);"#,
        ["Ada", "42"]
    };

    tuple_double_int => {
        r#"var (rate,count)=(2.5,4); Console.WriteLine(rate*count);"#,
        ["10"]
    };

    tuple_bool_pair => {
        r#"var (on,flag)=(true,false); Console.WriteLine(on); Console.WriteLine(flag);"#,
        ["True", "False"]
    };

    tuple_discard_second => {
        r#"var (_,y)=(99,3); Console.WriteLine(y);"#,
        ["3"]
    };

    tuple_discard_first => {
        r#"var (x,_)=(7,8); Console.WriteLine(x);"#,
        ["7"]
    };

    nested_tuple_flatten => {
        r#"var ((a,b),c)=((1,2),3); Console.WriteLine(a+b+c);"#,
        ["6"]
    };

    nested_tuple_three_level => {
        r#"var ((a,b),(c,d))=((1,2),(3,4)); Console.WriteLine(a+b+c+d);"#,
        ["10"]
    };

    tuple_from_method => {
        r#"System.ValueTuple<int,int> Pair()=>(4,5); var (x,y)=Pair(); Console.WriteLine(x+y);"#,
        ["9"]
    };

    tuple_from_local_function => {
        r#"System.ValueTuple<int,int> Twice(int n)=>(n,n); var (a,b)=Twice(6); Console.WriteLine(a*b);"#,
        ["36"]
    };

    tuple_swap => {
        r#"int a=1,b=2; (a,b)=(b,a); Console.WriteLine(a); Console.WriteLine(b);"#,
        ["2", "1"]
    };

    tuple_foreach_array => {
        r#"var pairs=new[]{(1,2),(3,4)}; int sum=0; foreach(var (x,y) in pairs) sum+=x+y; Console.WriteLine(sum);"#,
        ["10"]
    };

    tuple_long_values => {
        r#"var (lo,hi)=(10000000000L,5L); Console.WriteLine(lo+hi);"#,
        ["10000000005"]
    };

    tuple_char_int => {
        r#"var (ch,n)=('A',1); Console.WriteLine(ch); Console.WriteLine(n);"#,
        ["A", "1"]
    };

    tuple_decimal => {
        r#"var (a,b)=(1.5m,2.5m); Console.WriteLine(a+b);"#,
        ["4.0"]
    };

    tuple_named_elements => {
        r#"var t=(X:2,Y:3); var (x,y)=t; Console.WriteLine(x+y);"#,
        ["5"]
    };

    record_deconstruct_two => {
        r#"record Point(int X,int Y); var (x,y)=new Point(3,4); Console.WriteLine(x); Console.WriteLine(y);"#,
        ["3", "4"]
    };

    record_deconstruct_three => {
        r#"record Triple(int A,int B,int C); var (a,b,c)=new Triple(1,2,3); Console.WriteLine(a+b+c);"#,
        ["6"]
    };

    record_deconstruct_foreach => {
        r#"record Point(int X,int Y); var pts=new[]{new Point(1,2),new Point(3,4)}; int sum=0; foreach(var (x,y) in pts) sum+=x+y; Console.WriteLine(sum);"#,
        ["10"]
    };

    record_struct_deconstruct => {
        r#"record struct Vec(int X,int Y); var (x,y)=new Vec(8,1); Console.WriteLine(x-y);"#,
        ["7"]
    };

    readonly_record_struct_deconstruct => {
        r#"readonly record struct Pair(int A,int B); var (a,b)=new Pair(2,5); Console.WriteLine(a*b);"#,
        ["10"]
    };

    record_deconstruct_after_with => {
        r#"record Pair(int A,int B); var q=(new Pair(1,2)) with{A=9}; var (a,b)=q; Console.WriteLine(a); Console.WriteLine(b);"#,
        ["9", "2"]
    };

    record_deconstruct_string => {
        r#"record Tag(string Name); var (name)=new Tag("z"); Console.WriteLine(name);"#,
        ["z"]
    };

    record_deconstruct_mixed => {
        r#"record Mix(int N,string S); var (n,s)=new Mix(7,"x"); Console.WriteLine(n); Console.WriteLine(s);"#,
        ["7", "x"]
    };

    custom_class_deconstruct_two => {
        r#"class Size{public int W,H; public void Deconstruct(out int w,out int h){w=W;h=H;}} var (w,h)=new Size{W=3,H=4}; Console.WriteLine(w+h);"#,
        ["7"]
    };

    custom_class_deconstruct_three => {
        r#"class Box{public int A,B,C; public void Deconstruct(out int a,out int b,out int c){a=A;b=B;c=C;}} var (a,b,c)=new Box{A=1,B=2,C=3}; Console.WriteLine(a+b+c);"#,
        ["6"]
    };

    custom_struct_deconstruct => {
        r#"struct Pair{public int X,Y; public void Deconstruct(out int x,out int y){x=X;y=Y;}} var (x,y)=new Pair{X=4,Y=6}; Console.WriteLine(x*y);"#,
        ["24"]
    };

    custom_class_deconstruct_single => {
        r#"class Wrap{public int V; public void Deconstruct(out int v){v=V;}} var (v)=new Wrap{V=11}; Console.WriteLine(v);"#,
        ["11"]
    };

    custom_deconstruct_foreach_list => {
        r#"class Pair{public int A,B; public void Deconstruct(out int a,out int b){a=A;b=B;}} var list=new System.Collections.Generic.List<Pair>{new Pair{A=1,B=2},new Pair{A=3,B=4}}; int sum=0; foreach(var (a,b) in list) sum+=a+b; Console.WriteLine(sum);"#,
        ["10"]
    };

    deconstruct_to_existing_locals => {
        r#"class Pair{public int A,B; public void Deconstruct(out int a,out int b){a=A;b=B;}} var target=new Pair{A=5,B=6}; int x,y; (x,y)=target; Console.WriteLine(x+y);"#,
        ["11"]
    };

    record_deconstruct_to_locals => {
        r#"record R(int A,int B); var r=new R(2,3); int x,y; (x,y)=r; Console.WriteLine(x*y);"#,
        ["6"]
    };

    deconstruct_nested_record_field => {
        r#"record Inner(int N); record Outer(Inner I); var (n)=new Outer(new Inner(9)); Console.WriteLine(n);"#,
        ["9"]
    };

    deconstruct_derived_record => {
        r#"record Animal(string Name); record Dog(string Name,int Age):Animal(Name); var (name,age)=new Dog("Rex",4); Console.WriteLine(name); Console.WriteLine(age);"#,
        ["Rex", "4"]
    };

    deconstruct_sequential => {
        r#"var (a,b)=(1,2); var (c,d)=(a+b,b); Console.WriteLine(c); Console.WriteLine(d);"#,
        ["3", "2"]
    };

    deconstruct_tuple_condition => {
        r#"var t=(2,5); var (x,y)=t; Console.WriteLine(x<y);"#,
        ["True"]
    };

    deconstruct_record_sum => {
        r#"record V(int A,int B,int C); var (a,b,c)=new V(1,2,3); Console.WriteLine(a+b+c);"#,
        ["6"]
    };

    deconstruct_bool_record => {
        r#"record Flag(bool On); var (on)=new Flag(true); Console.WriteLine(on);"#,
        ["True"]
    };

    deconstruct_char_record => {
        r#"record Sym(char C); var (c)=new Sym('Q'); Console.WriteLine(c);"#,
        ["Q"]
    };

    deconstruct_byte_record => {
        r#"record ByteVal(byte B); var (b)=new ByteVal(255); Console.WriteLine(b);"#,
        ["255"]
    };

    deconstruct_negative_tuple => {
        r#"var (a,b)=(-3,-7); Console.WriteLine(a+b);"#,
        ["-10"]
    };

    deconstruct_zero_tuple => {
        r#"var (a,b)=(0,0); Console.WriteLine(a==b);"#,
        ["True"]
    };

    deconstruct_existing_then_sum => {
        r#"int s=0,t=0; (s,t)=(4,6); Console.WriteLine(s+t);"#,
        ["10"]
    };

    deconstruct_record_four_fields => {
        r#"record Quad(int A,int B,int C,int D); var (a,b,c,d)=new Quad(1,2,3,4); Console.WriteLine(d);"#,
        ["4"]
    };

    custom_deconstruct_discard => {
        r#"class Pair{public int A,B; public void Deconstruct(out int a,out int b){a=A;b=B;}} var (_,b)=new Pair{A=9,B=2}; Console.WriteLine(b);"#,
        ["2"]
    };

    deconstruct_tuple_from_array => {
        r#"var arr=new[]{(1,2),(3,4)}; var (x,y)=arr[1]; Console.WriteLine(x+y);"#,
        ["7"]
    };

    deconstruct_record_enum => {
        r#"enum Mode{Off,On} record State(Mode M); var (m)=new State(Mode.On); Console.WriteLine(m);"#,
        ["On"]
    };

    deconstruct_nullable_record => {
        r#"record Maybe(int? N); var (n)=new Maybe(5); Console.WriteLine(n);"#,
        ["5"]
    };

    deconstruct_double_record => {
        r#"record Rate(double V); var (v)=new Rate(3.5); Console.WriteLine(v);"#,
        ["3.5"]
    };

    deconstruct_decimal_record => {
        r#"record Money(decimal A); var (a)=new Money(9.99m); Console.WriteLine(a);"#,
        ["9.99"]
    };

    deconstruct_tuple_to_method => {
        r#"void Sum(int a,int b){Console.WriteLine(a+b);} var (x,y)=(2,3); Sum(x,y);"#,
        ["5"]
    };

    deconstruct_record_twice => {
        r#"record Pair(int A,int B); var p=new Pair(1,2); var (a,b)=p; var (c,d)=p; Console.WriteLine(a+c); Console.WriteLine(b+d);"#,
        ["2", "4"]
    };
}
