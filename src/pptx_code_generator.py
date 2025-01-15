from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI

class PPTXCodeGenerator:
    def __init__(self, llm: ChatOpenAI):
        self.llm = llm
        
    def run(self, slide_contents: str) -> str:
        # プロンプトを定義
        prompt = ChatPromptTemplate.from_messages(
            [
                (
                    "system",
                    "python-pptxモジュールを用いてプレゼンテーション資料のスライドを自動生成する専門家です。"
                ),
                (
                    "human",
                    "以下のスライドの内容を生成するためのpythonコードを生成してください。\n\n"
                    "スライドの内容:\n{slide_contents}\n\n"
                    "以下のpptxファイルを読み込み、テンプレートとして使用してください。\n"
                    "/workspace/input/template.pptx\n"
                    "テンプレートのレイアウト情報は以下を参照してください。記載にないレイアウト番号やプレースホルダー番号は決して使用しないでください。\n"
                    "    タイトルスライド: slide_layouts[2]\n"
                    "        placeholder_format.idx:\n"
                    "            0: 会社名など\n"
                    "            10: 発表タイトル\n"
                    "            11: サブタイトル・日付など\n"
                    "            12: 発表者名など\n"
                    "    一般スライド: slide_layouts[0]\n"
                    "        placeholder_format.idx:\n"
                    "            0: スライドタイトル\n"
                    "            1: 内容\n"
                    "作成したパワーポイントは/workspace/output内に出力されるようにしてください。\n\n"
                    "ルール:\n"
                    "- 【重要】必ずpython-pptxモジュールを使用したpythonコードのみを出力してください。\n"
                    "- 使用が許可されているのは、テキスト、図形、表のみです。\n"
                    "- テキスト以外の要素（図形および表）を使用してほしい箇所には、その旨が明記されています。\n"
                    "- 画像や動画は使用できません。絶対に画像や動画を使用しないでください。\n"
                    "- '---next---' はスライド番号を進める合図です。このタイミングで新たなスライドを追加してください。\n\n"
                )
            ]
        )
        # スライド生成のためのチェーンを作成
        chain = prompt | self.llm | StrOutputParser()
        # スライド生成のコードを生成
        return chain.invoke({"slide_contents": slide_contents})