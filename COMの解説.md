---
marp: true
theme: default
---

# C++勉強会
## 第1回 ~ COM ~

---

# 本日伝えたいこと

1.COM（Component Object Model）はWindows上で部品（コンポーネント）を作って、言語やプロセスを超えて再利用できる仕組みです。

2.「全言語で共通に扱える型」としてCOMの独自型 BSTR / VARIANT / SAFEARRAY /HRESULTが定義されています。

3.COMの独自型を扱うときメモリ管理をする必要があるため、ラッパーのクラス（CComBSTR、CComVariant、CComSafeArray）を使用すべき。

4.COMは古い技術のため、.NETなどの新しい技術を今後のアプリ開発では使用すべきです。

---

# COMが登場する前

COMが出現する前は、プロセス間(.exe)通信のハードルが高かった。
例）C++、VB、C# などで書かれたアプリを同じ方法で呼び出すのが困難。
例）共有メモリやファイルベースはデータ整合性や同期の管理が煩雑。

↓　そこで登場したのが

COM（Component Object Model）：Windows上で部品（コンポーネント）を作って、言語やプロセスを超えて再利用できる仕組み

---

# COMが登場した後

プロセス間(.exe)通信のハードル下がり、2000年代で主流になった。

## COMの特徴

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
・COMはC++、VB, Delphi, .NET, スクリプト言語（VBScript, JavaScript）など 複数の言語から利用される前提。

・しかし、各言語で「型の表現方法」が違う。
C++: std::string / int / bool
VB: String / Integer / Boolean
JavaScript: string / number / boolean

・この差を吸収するために「全言語で共通に扱える型」として BSTR / VARIANT / SAFEARRAY /HRESULTなどを定義。

---

# BSTR型
BSTR は COM で利用される一般的な文字列の形式です。

BSTR は基本的にワイドキャラクタ文字列(C++では先頭にL記述)として定義されています。
その手前にデータ領域のバイト数を示すバイト数（４バイト）と、 終端に (00 00) という２バイトのデータが割り当てられます。

<img src=".\Img\BSTR型.png" width="600">

---

# BSTR型
## 使い方

<NG>

データバイト数を表すバイト数が割り当てられてないので、正しい BSTR型ではない。

~~~ C++
BSTR bstr = L"Hello";
~~~

<OK>
BSTR型の割り当てには SysAllocString、解放には SysFreeString を用います。これをしないとメモリリークする。

~~~ C++
// メモリ確保
BSTR bstr = SysAllocString( L"Hello" );
// 解放
SysFreeString( bstr );
~~~

---
# BSTR型
## ラッパークラス

CComBSTRを使用すると、コンストラクタでSysAllocString、デストラクタでSysFreeStringが自動的に呼ばれる。

~~~ C++
{
    CComBSTR bstr1(L"Hello");  // ← 内部で SysAllocStringを実行
} // ← スコープを抜けると ~CComBSTR() が呼ばれ SysFreeStringを実行
~~~

ラッパークラスはメモリの自動解放だけでなく、文字列結合、比較などを簡単にできる

---
# BSTR型
## ラッパークラス
### 文字列結合

<ラッパークラスを使用しない場合>
BSTR型の文字列結合だけでも10行程度必要。

~~~ C++
BSTR str = SysAllocString(L"Hello");
BSTR add = SysAllocString(L" World");

// 新しい長さ分を確保
UINT len1 = SysStringLen(str);
UINT len2 = SysStringLen(add);
BSTR newStr = SysAllocStringLen(NULL, len1 + len2);

// 値コピー
memcpy(newStr, str, len1 * sizeof(OLECHAR));
memcpy(newStr + len1, add, len2 * sizeof(OLECHAR));

// 後処理
SysFreeString(str);
SysFreeString(add);
~~~

<style scoped>
section {
  font-size: 160%;   /* 全体の文字サイズ */
}
</style>

---

# BSTR型
## ラッパークラス
### 文字列結合

<ラッパークラスを使用する場合>
3行で実装が可能。

~~~ C++
    CComBSTR str(L"Hello");
    CComBSTR add(L" World");

    str += add;

~~~

結論：積極的にラッパークラスを使用すべき。

---

# VARIANT型

VARIANT型は「なんでも入る箱」（汎用データ型）
COMで数値・文字列・配列などをやり取りするときに使う

データ構造：データ型（vt）＋ データ(共用対)

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

<style scoped>
section {
  font-size: 160%;   /* 全体の文字サイズ */
}
</style>

---
# VARIANT型
## 使い方

VARIANT型を使うにはVariantInit、解放にはVariantClearを用います。
値代入のときには型情報＆値を入れる必要がある。

~~~ C++
//初期化
VARIANT var;
VariantInit(&var); // vt = VT_EMPTY、データ部分を0クリア

// 数値を入れる場合
var.vt = VT_I4;
var.lVal = 42;

// 取り出すときは型を見て取り出す
if (var.vt == VT_I4) {
    long value = var.lVal;  // 42 が取り出せる
}

// 解放処理
VariantClear(&var); // vtを見て正しく解放。vt=VT_EMPTY に戻す
~~~

<style scoped>
section {
  font-size: 150%;   /* 全体の文字サイズ */
}
</style>

---

# VARIANT型
## ラッパークラス
CComVariant型は自動でVARIANTInit、VariantClearをしてくれる。
代入された値に合わせて型情報も変更してくれる。

~~~ C++
    // 初期化（内部でVARIANTInitが実行）
    CComVariant var;

    // 数値を代入
    var = 42;  // 自動的に vt = VT_I4 になる

    // 取り出しはVARIANT型と同じ
    if (var.vt == VT_I4) {
        long value = var.lVal;  // 42 が取り出せる
    }

    // 解放処理は不要(内部でVariantClearが実行)
~~~

CComVariant型も便利なメンバー関数（Copyなど）があるが説明は省略する。

<style scoped>
section {
  font-size: 160%;   /* 全体の文字サイズ */
}
</style>

---
# SafeArray型

C/C++の生配列（`int arr[10]など`）は以下の問題がある。

- 長さ情報がない
- 範囲外アクセスが検出できない
- 言語ごとに表現が違う

そこで COM では、**配列を安全に受け渡しするための標準コンテナ**として `SAFEARRAY型` を使用する

---

# SafeArray型

`SAFEARRAY` は構造体で、次の情報を持っています：

- **配列の次元数（1次元・2次元…）**
- **各次元の下限・上限**
- **要素の型（int, double, BSTR, VARIANTなど）**
- **データへのポインタ**

---
# SafeArray型
## 使い方（SafeArrayPutElement / SafeArrayGetElement）

~~~ C++
// 要素の情報
SAFEARRAYBOUND bound;
bound.lLbound = 0;        // 下限（配列の開始インデックス）
bound.cElements = 5;      // 要素数

// SAFEARRAY生成
SAFEARRAY* psa = SafeArrayCreate(VT_I4, 1, &bound); // int32 の1次元配列を作成

// 値格納
long index = 0;
int value = 123;
SafeArrayPutElement(psa, &index, &value);

// 値取得
int outValue;
SafeArrayGetElement(psa, &index, &outValue);

// 後処理
SafeArrayDestroy(psa);
~~~

SafeArrayPutElement / SafeArrayGetElement は単純な配列アクセスではなく 関数呼び出しとチェックを伴う ため、処理が遅くなります。

<style scoped>
section {
  font-size: 150%;   /* 全体の文字サイズ */
}
</style>

---
# SafeArray型
## 使い方（SafeArrayAccessData/SafeArrayUnaccessData）
SafeArrayPutElement / SafeArrayGetElementより高速に処理したい場合はSafeArrayAccessData/SafeArrayUnaccessDataを使用してポインタに直接アクセスする。
何千、万件というデータにアクセスする場合はこちらを使用したほうがよい。

~~~ C++
// 要素の情報
SAFEARRAYBOUND bound;
bound.lLbound = 0;        // 下限（配列の開始インデックス）
bound.cElements = 5;      // 要素数

// SAFEARRAY生成
SAFEARRAY* psa = SafeArrayCreate(VT_I4, 1, &bound); // int32 の1次元配列を作成

long* pData = nullptr;
SafeArrayAccessData(psa, (void**)&pData);  // 生ポインタを取得&配列のロック

// 値格納
pData[0] = 123;

// 値取得
long outValue = pData[0];

SafeArrayUnaccessData(psa);                 // 解放

// 後処理
SafeArrayDestroy(psa);
~~~

<style scoped>
section {
  font-size: 130%;   /* 全体の文字サイズ */
}
</style>

---

# CComSafeArray型
## ラッパークラス
CComSafeArray型はコンストラクタでSafeArrayCreate、デストラクタでSafeArrayDestroyを自動で処理してくれる。
SAFEARRAYBOUND（配列情報）の作成も不要。

~~~ C++
    // CComSafeArray<int> で int32 配列を作成（下限 0、要素数 5）
    CComSafeArray<int> arr(5);

    // 要素をセット
    arr.SetAt(0, 123);　　//SafeArrayPutElementの代わり

    // 要素を取得
    int outValue = arr.GetAt(0);   // SafeArrayGetElementの代わり

    // 解放処理は自動
~~~

SafeArrayAccessDataは同様に使用可能。

<style scoped>
section {
  font-size: 170%;   /* 全体の文字サイズ */
}
</style>

---

# COMは現在主流でない理由

- 煩雑なメモリ管理
ラッパークラスで緩和はされているが、VARIANT や BSTR の初期化・解放を正確に行う必要がある。

- 非クロスプラットフォーム
COMはWindowsの技術。非クロスプラットフォーム(Linuxなどでは使用不可)。

- 学習コストが高い
レジストリ登録、独自の型（BSTR、VARIANT、SAFEARRAY）

- 後継技術の登場
COM の代替として、より簡単で安全にコンポーネント化・オブジェクト連携できる.NET,gRPCなどの仕組みに押されていった

➡COMは古い技術なので、今からアプリ開発をするときは使用を避けるべき。

---

# 本日伝えたいこと

1.COM（Component Object Model）はWindows上で部品（コンポーネント）を作って、言語やプロセスを超えて再利用できる仕組みです。

2.「全言語で共通に扱える型」としてCOMの独自型 BSTR / VARIANT / SAFEARRAY /HRESULTが定義されています。

3.COMの独自型を扱うときメモリ管理をする必要があるため、ラッパーのクラス（CComBSTR、CComVariant、CComSafeArray）を使用すべき。

4.COMは古い技術のため、.NETなどの新しい技術を今後のアプリ開発では使用すべきです。
