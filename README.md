# worcry-automations

WorCry(ワークライ)社内の業務自動化スクリプト置き場です。

## scripts/daily_shukatu_news.py

大学3〜4年生の就活本番期向けに、前日公開の経済・ビジネス・社会ニュースを
Claude(Anthropic API + web_searchツール)で収集し、1本を選定してSlackの
「就活ニュース」チャンネルに自動投稿します。

### 必須環境変数

- `ANTHROPIC_API_KEY`
- `SLACK_WEBHOOK_URL`

### 重複防止

`scripts/post_history.log` に投稿履歴を追記し、過去テーマとの重複を回避します。

## .github/workflows/daily-shukatu-news.yml

平日(JST) 8:00 = UTC 前日 23:00 に GitHub Actions で自動実行します。
手動実行は Actions タブの「Daily Shukatu News」→「Run workflow」から可能です。

Secrets に `ANTHROPIC_API_KEY` と `SLACK_WEBHOOK_URL` が設定されている必要があります。
