---
marp: true
theme: default
---

# C++勉強会
## 第1回 ~ COM ~

---

# 前提

前提：COMは古い技術なので、今からアプリ開発をするときは使用を避けるべき。
世の中にはCOMに代わる新しい技術が誕生している。

ORiN2がCOMを利用しているので、今回は仕方なく学習しているだけ。

---

# COMが登場する前

COMが出現する前は、プロセス間(.exe)通信がのハードルが高かった。
例）C++、VB、C# などで書かれたアプリを同じ方法で呼び出すのが困難。
例）共有メモリやファイルベースはデータ整合性や同期の管理が煩雑。

↓　そこで登場したのが

COM（Component Object Model）：Windows上で部品（コンポーネント）を作って、言語やプロセスを超えて再利用できる仕組み

---

# COMが登場した後

COMが登場したことで、プロセス間(.exe)通信のハードル下がり、2000年代で主流になった。

- **言語に依存しない**
    
    C++、VB、C#、スクリプトなど、どの言語からでも使える。
    
- **バイナリ互換性**
    
    DLLやEXEをコンパイルし直さなくてもやり取りできる。
    
- **インターフェース指向**
    
    実装は隠して、「こういう機能を持っています」という契約（インターフェース）だけ公開。

---

ORiNはCOMに依存しているので、様々な言語のアプリから呼び出しが可能。
CAO.exeが他アプリから再利用されている。

<img src=".\Img\ORiNの説明.png">


---

# なぜCOMの独自型が必要なのか？
・COMはC++専用ではなく、VB, Delphi, .NET, スクリプト言語（VBScript, JavaScript）など 複数の言語から利用される前提。

・しかし、各言語で「型の表現方法」が違う。
C++: std::string / int / bool
VB: String / Integer / Boolean
JavaScript: string / number / boolean

・この差を吸収するために「全言語で共通に扱える型」として BSTR / VARIANT / SAFEARRAY /HRESULTなどを定義。

・ただし、この共通型を使う

---

# BSTR型
BSTR は COM で利用される一般的な文字列の形式です。

BSTR は基本的にワイドキャラクタ文字列として定義されていますが、その手前にデータ領域のバイト数を示すバイト数（４バイト）と、 終端に (00 00) という２バイトのデータが割り当てられます。


<img src=".\Img\BSTR型.png">

---

<NG例>
データバイト数を表すバイト数が割り当てられてないので、正しい BSTR型ではない。

~~~ C++
BSTR bstr = L"Hello";
~~~

<OK例>
BSTR型の割り当てには SysAllocString、解放には SysFreeString を用います。

~~~ C++
// メモリ確保
BSTR bstr = SysAllocString( L"Hello" );
// 解放
SysFreeString( bstr );
~~~

---

## BSTRラッパークラス

コンストラクタで自動的にSysAllocString
デストラクタで自動的にSysFreeString

~~~ C++
BSTR bstr1 = CComBSTR(L"Hello");
// 解放処理不要。スコープ抜けたら自動的に開放。
~~~

---

## BSTRラッパークラス

<ラッパークラスを使用しない場合>
BSTR型の文字列結合だけでも10行程度必要。

~~~ C++
BSTR str = SysAllocString(L"Hello");
BSTR add = SysAllocString(L" World");

// 新しい長さ分を確保
UINT len1 = SysStringLen(str);
UINT len2 = SysStringLen(add);
BSTR newStr = SysAllocStringLen(NULL, len1 + len2);

memcpy(newStr, str, len1 * sizeof(OLECHAR));
memcpy(newStr + len1, add, len2 * sizeof(OLECHAR));

// 後処理
SysFreeString(str);
SysFreeString(add);
~~~

---

## BSTRラッパークラス

<ラッパークラスを使用する場合>
3行で実装が可能。

~~~ C++
    CComBSTR str(L"Hello");
    CComBSTR add(L" World");

    str += add;

~~~

積極的にラッパークラスを使用すべき。

---

# VARIANT型
一言でいうと「なんでも入る箱」（汎用データ型）
COMで数値・文字列・オブジェクト・配列などをやり取りするときに使う

データ：型タグ（vt）＋ データのunion

~~~ C++
typedef struct tagVARIANT {
    VARTYPE vt;   // 値の型を表す
    WORD    wReserved1;
    WORD    wReserved2;
    WORD    wReserved3;
    union {
    LONG        lVal;
    BSTR        bstrVal;
    SAFEARRAY*  parray;
    // ... 他にも多数
    };
} VARIANT;
~~~

---
<使い方>

VariantInit、VariantClear、型情報＆値を正確に入れる必要がある。

~~~ C++
//初期化
VARIANT var;
VariantInit(&var); // vt = VT_EMPTY で初期化

// 数値を入れる場合
var.vt = VT_I4;
var.lVal = 42;

// 文字列を入れる場合
var.vt = VT_BSTR;
var.bstrVal = SysAllocString(L"Hello"); // 解放はVariantClearに任せる

// 解放処理
VariantClear(&var); // vtを見て正しく解放してくれる
~~~

---

## CComVARIANT型
VARIANT型のラッパークラス
自動でVARIANTInit、VariantClearをしてくれる。
代入された値に合わせて型情報も変更してくれる。

~~~ C++
    // 初期化（VARIANTInit不要）
    CComVariant var;

    // 数値を代入
    var = 42;  // 自動的に vt = VT_I4 になる

    // 文字列を代入
    var = L"Hello";  // 自動的に vt = VT_BSTR になり、内部で SysAllocString される

    // 解放処理も不要

~~~

---
## SafeArray型

### 一言でいうと

C/C++の生配列（`int arr[10]`）は

- 長さ情報がない
- 範囲外アクセスが検出できない
- 言語ごとに表現が違う

という問題があります。

そこで COM では、**配列を安全に受け渡しするための標準コンテナ**として `SAFEARRAY` を導入しました。

---

### 2. **構造**

`SAFEARRAY` は構造体で、次の情報を持っています：

- **配列の次元数（1次元・2次元…）**
- **各次元の下限・上限**
- **要素の型（int, double, BSTR, VARIANTなど）**
- **データへのポインタ**

つまり **「メタ情報 + データ」** がひとまとめになっています。

---

~~~ C++
SAFEARRAYBOUND bound[1];
bound[0].lLbound = 0;        // 下限（配列の開始インデックス）
bound[0].cElements = 5;      // 要素数

// SAFEARRAY生成
SAFEARRAY* psa = SafeArrayCreate(VT_I4, 1, bound); // int32 の配列を作成

// Put
long index = 0;
int value = 123;
SafeArrayPutElement(psa, &index, &value);

// Get
int outValue;
SafeArrayGetElement(psa, &index, &outValue);

// 後処理
SafeArrayDestroy(psa);
~~~

---

## CComSafeArray型
SafeArray型のラッパークラス
自動でVARIANTInit、VariantClearをしてくれる。
代入された値に合わせて型情報も変更してくれる。

~~~ C++
    // CComSafeArray<int> で int32 配列を作成（下限 0、要素数 5）
    CComSafeArray<int> arr(5);

    // 要素をセット
    arr.SetAt(0, 123);

    // 要素を取得
    int outValue = arr.GetAt(0);

    // 解放処理は自動
~~~


---

# COMは現在主流でない理由

- 複雑で使いにくい
メモリ管理が手作業（`AddRef` / `Release`、`SysAllocString` / `SysFreeString` ）
ラッパーは出たもののC#のようにGCでOSがメモリ管理してくれるような機能はない。
- 学習コストが高い
レジストリ登録、独自の型（BSTR、VARIANT、SAFEARRAY）
- 非クロスプラットフォーム
COMはWindowsの技術。非クロスプラットフォーム。
- 後継技術の登場
.NETなどのOSを超えて使える仕組みに押されていった

➡COMは古い技術なので、今からアプリ開発をするときは使用を避けるべき。