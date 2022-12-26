using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Produire;

// アセンブリに関する一般情報は以下の属性セットをとおして制御されます。
// アセンブリに関連付けられている情報を変更するには、
// これらの属性値を変更してください。
[assembly: AssemblyTitle("Susieホスト プラグイン")]
[assembly: AssemblyDescription("プロデルからSusieプラグインを利用できます")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("Susieホスト プラグイン")]
[assembly: AssemblyCopyright("Copyright(C) 2007-2022 utopiat.net")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

[assembly: PluginName("Susieプラグイン")]
[assembly: PluginDescription("プロデルからジョイステックやゲームパッドの情報を取得できます")]
[assembly: PluginVersion("1.0")]
[assembly: PluginAcceptVersion("1.0")]

// ComVisible を false に設定すると、このアセンブリ内の型は COM コンポーネントから
// 参照できなくなります。COM からこのアセンブリ内の型にアクセスする必要がある場合は、
// その型の ComVisible 属性を true に設定してください。
[assembly: ComVisible(false)]

// このプロジェクトが COM に公開される場合、次の GUID が typelib の ID になります
[assembly: Guid("82a12ea7-ae92-470c-847a-15c7abaeb756")]

// アセンブリのバージョン情報は、以下の 4 つの値で構成されています:
//
//      メジャー バージョン
//      マイナー バージョン
//      ビルド番号
//      リビジョン
//
// すべての値を指定するか、次を使用してビルド番号とリビジョン番号を既定に設定できます
// 既定値にすることができます:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
