import argparse
from langchain_openai import ChatOpenAI

from pptx_agent import PPTXAgent

def main():
    # コマンドライン引数のパーサーを作成
    parser = argparse.ArgumentParser(
        description="ユーザーがインプットしたテキストを基にスライドを生成するPythonファイルを出力します"
    )
    # "file"引数を追加
    parser.add_argument(
        "--file",
        type=str,
        help="プレゼンテーションの元になるテキストファイルのパス(.txt/.docx/.md)"
    )
    # コマンドライン引数を解析
    args = parser.parse_args()
    
    # テキストの取得
    filepath = args.file
    if filepath.endswith((".txt", ".docx", ".md")):
        with open(filepath, "r") as f:
            user_request = f.read()
    else:
        raise ValueError("ファイル形式がサポートされていません")
    
    # ChatOpenAIモデルを初期化
    llm = ChatOpenAI(model="gpt-4o", temperature=0.0)
    # PPTXAgentを初期化
    agent = PPTXAgent(llm=llm)
    # エージェントを実行して最終的な出力を取得
    final_output = agent.run(user_request=user_request)
    final_output = final_output.split("```python\n")[-1].split("```")[0]
    # 出力をファイルに保存
    with open("/workspace/output/create_pptx.py", "w") as f:
        f.write(final_output)
        
    print("DONE.")
    
if __name__ == "__main__":
    main()