---
languages:
  - C#
products:
  - windows-wpf
page_type: sample
name: "WPF Custom IDE"
description: "메모장으로 WPF용 IDE를 만든 군대 선임의 작품을 클론코딩 해보았습니다."
---

# WPF Custom IDE

군대 인트라넷에는 윈도우 10을 설치하면서 깔린 .Net 프레임워크 말고는 개발 도구가 없습니다. (특기학교 컴퓨터나 개발대 사무실 등 언제나 예외는 있습니다.) 그런 열악한 환경에서 사무실의 업무 편의를 위해 메모장으로 윈도우 프로그램(WPF)을 만든 선임이 있었습니다. 그 분은 IDE까지 손수 만들어서 그 IDE를 이용해 프로그램을 만들었습니다.
이 프로젝트는, 아무도 시키지 않았지만, 그 프로그램을 유지보수 해보기 위해 우선 IDE부터 제 스스로 따라 만들어 보고자 시작한 프로젝트입니다.
<br>구체적인 진행과정은 [블로그](https://chinpa.tistory.com)에 적고 있습니다.

## How it works

IDE라고 이름은 거창하지만 사실 구조를 뜯어보면 단순합니다.
윈도우 운영체제에 기본으로 존재하는 MSBuild 라는 빌드 도구를 백그라운드 Process로 호출해 WPF를 빌드하면 컴파일러+빌더가 벌써 IDE에 들어왔습니다.
따라서 이 IDE 프로그램은 사실상 텍스트 편집기 기능을 가능한 VSCode와 유사하게 만드는데 그 목적을 둘 것입니다.

### To Do List

- ~~기본 윈도우 프로그램 형식 만들기 (완)~~
- 프로그램 레이아웃 디자인
- 텍스트 편집기능 추가
- ~~빌드~~, 실행 기능 추가 (진행중)
- Line Number 기능 추가

### Reference
- [왕초보개발자](https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=tablesetter&logNo=20130207390)
- [고급WPF영역](https://docs.microsoft.com/ko-kr/dotnet/desktop/wpf/advanced/flow-document-overview?view=netframeworkdesktop-4.8)
- [컨트롤DOC](https://docs.microsoft.com/en-us/dotnet/desktop/wpf/controls/richtextbox-overview?view=netframeworkdesktop-4.8)
- [Syntax Highlight Example](https://www.c-sharpcorner.com/article/syntax-highlighting-in-rich-textbox-control-part-1/)
- [Syntax Highlight Issue Stack Overflow](https://stackoverflow.com/questions/17971300/wpf-richtextbox-syntax-highlighting-issue)


