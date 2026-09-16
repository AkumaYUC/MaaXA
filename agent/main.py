import os
import sys

#便携 Python（embeddable）带 ._pth，不会自动把脚本所在目录加进 sys.path，
#同级模块 import 会失败；这里显式补上（开发环境用 venv 时同样无副作用）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from maa.agent.agent_server import AgentServer
from maa.toolkit import Toolkit

import my_action
import common
import transfer_actions


def main():
    Toolkit.init_option("./")

    if len(sys.argv) < 2:
        print("Usage: python main.py <socket_id>")
        print("socket_id is provided by AgentIdentifier.")
        sys.exit(1)
        
    socket_id = sys.argv[-1]

    AgentServer.start_up(socket_id)
    AgentServer.join()
    AgentServer.shut_down()


if __name__ == "__main__":
    main()
