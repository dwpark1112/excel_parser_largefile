# How to load a large xlsx file with Apache POI?

Java 어플리케이션에서 엑셀 파일을 읽고, 무언가 처리하려면 Apache POI 라이브러리를 사용하면 된다. 익히기 쉽고 간단한데 하나 문제가 있다면, 메모리를 무지막지하게 소비해버린다는 것이다. 게다가 느리다. 마치 DOM parser 처럼 모든 row, column을 전부 메모리에 적재하는 것 같다.

```java
// 느린 부분
Workbook workbook = new XSSFWorkbook( inputStream );
```

회사에서 관리하는 병원 중에는 우리 서비스 말고도 외부 서비스를 사용하고 있는 병원이 있는데, 그 쪽에서는 우리쪽 통계 시스템에 모든 통계 정보를 저장하고 집계하길 원했다. 외부 서비스의 통계 정보는 엑셀 파일로 나오는데, 나오는 시간도 상당히 오래걸리지만 그 양도 만만치가 않았다. 약 80만건의 엑셀 데이터 정보;;

행정 업무 보는 분들이 엑셀파일을 올릴 수 있도록 만들긴 했는데, 릴리즈할 자신이 없다. 5만건의 row, 14개의 column으로 구성된 엑셀 파일 하나만 올려도 필시 서버가 OutOfMemory 상태가 될 것이기에...

POI에는 Streaming으로 리딩하거나 일정 단위로 끊어서 메모리에 적재할 수 있는 구현체가 없는가 찾아보다가 어느 훌륭하신 분이 만든 Streaming Reader를 발견했다. <https://github.com/monitorjbl/excel-streaming-reader>

테스트로 잠깐 동일 파일 실행 비교를 해보니깐, 이건 뭐 측정해보지 않아도 체감되는 성능이다.

## What

`난 엑셀 파일을 빠르게 읽고 싶어!`라는 사용자를 위한 라이브러리, Apache POI가 모든 데이터를 Workbook에 적재하는 것 대신, excel-streaming-reader는 streaming API를 감싸는 역할을 수행해서 POI를 사용했던 사용자도 쉽게 쓸 수 있게 만들어 두었다.

> You may access cells randomly within a row, as the entire row is cached. **However**, there is no way to randomly access rows. As this is a streaming implementation, only a small number of rows are kept in memory at any given time.

README에 적혀있던 내용인데, Row단위로만 메모리에 적재하고 SAX 파서처럼 순차적으로 읽어나가기 때문에, random하게 row를 찾아갈 순 없고 하나의 row내의 cell은 자유롭게 접근 가능하다는 것

그러니깐 혹시라도 row를 순차적으로 읽는것이 아닌 이리저리 이동하며 읽어야 한다면, 이 라이브러리는 적합하지 않다.

## HOW

Workbook을 생성하는 부분을 아래 처럼 StreamingReader의 빌더를 사용해서 구성하면 된다.
이 라이브러리는 comment가 없다. 근데 comment가 필요한것 같지도 않다.

```java
InputStream is = new FileInputStream(new File("/path/to/workbook.xlsx"));
Workbook workbook = StreamingReader.builder()
        .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
        .open(is);            // InputStream or File for XLSX file (required)
```

README도 깔끔하고, 설명과 주의사항을 읽다보니 이 분은 참 뛰어난 개발자 같다는 생각이 든다.

> This library will ONLY work with XLSX files. The older XLS format is not capable of being streamed.