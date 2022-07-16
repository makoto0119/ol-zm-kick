# ol-zm-kick
ol-zm-kick は、 outlook の今日の予定の中で、次に開始される zoom 会議の開始時間に zoom を起動するツールです。 zoom だけでなく、 teams の会議にも対応しています。 python で書かれていて、Windows で動作確認済みです。

## logic
ソースを見て下さい。そこそこ例外動作は入れてあります。

## Features
1. zoom と teams 会議に対応しています。 URL を作って起動するため、web ブラウザの画面が出ますが、仕様です。
2. zoom は、パスコード（パスワード）を分離したメールが来ても、合成して入力不要とします。
3. teams の会議にも対応しています（自動判別）。
4. 15分前から起動可能で、会議 30秒前までタイマーをかけて待機します。
5. 下記 hourglass という タイマーアプリを叩くようにしてあります。

## Requirement
ソースを持って行った人は、エラーが出たら個別に対応して下さい。

## Installation
python の実行環境があれば、ol-zm-kick.py を DL して使って下さい。
最大待機時間(2H)や、事前起動時間(30秒前)の設定など、必要ならソースを修正して下さい。
https://chris.dziemborowicz.com/apps/hourglass/ の タイマーアプリを併用すると便利です。不要なら該当行の削除をお願いします。

## Usage
動かすだけです。特に UI はありません。勝手に outlook を検出し、情報を取り、zoom や teams に URL でアクセスします。

## Note
outlook, zoom, teams の仕様が変わると、動かなくなる可能性はあります。
例えば、zoom は最初は「パスワード」というキーワードでメールが来ていましたが、現在は「パスコード」で来ます。
これには対応していますが、今後は分かりません。何かあれば、連絡下さい。

## Author
* 齋藤　誠
* 個人
* makoto0119@gmail.com

## License
自由に使って下さい
