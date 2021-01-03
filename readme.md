## 数据仓库项目 数据清洗
### 文件结构说明
- data -- 源数据和生成数据文件
    - source.xlsx -- 初始文件
    - movies.xlsx -- ExtractMovies.py运行后生成的电影文件
    - CleanedMovies.xlsx -- CleanData.py运行后生成的清理后的电影文件
    - Step1Result.xlsx -- DisposeDuplicates.py运行后生成的第一阶段清洗结果
    - imdbTest.xlsx -- imdb爬取结果
    - MergedData.xlsx -- JoinImdb.py运行结果
    - FinalMovies.xlsx -- MergeAndClean.py运行结果
    - Actors.xlsx -- SplitActors.py运行结果
    - Directors.xlsx -- SplitDirectors.py运行结果
- DataCleaning -- python包
    - ExtractMovies.py -- 提取电影
        - 输入：source.xlsx
        - 输出：movies.xlsx
        
    - CleanData.py -- 数据清理
        - 输入：movies.xlsx
        - 输出：CleanedMovies.xlsx
      
    - DisposeDuplicates.py -- 去重
        - 输入：CleanedMovies.xlsx
        - 输出：Step1Result.xlsx
      
    - JoinImdb.py -- 联合imdb数据
        - 输入：imdbTest.xlsx，Step1Result.xlsx
        - 输出：MergedData.xlsx
      
    - MergeAndClean.py -- 最后清洗
        - 输入：MergedData.xlsx
        - 输出：FinalMovies.xlsx
      
    - SplitActors.py -- 分离演员
        - 输入：FinalMovies.xlsx
        - 输出：Actors.xlsx
      
    - SplitDirectors.py -- 分离导演
        - 输入：FinalMovies.xlsx
        - 输出：Directors.xlsx
      
    - test.py -- 写代码时候顺手测试的代码
    
### 第一阶段数据清洗操作
1. 将数据整合之后的xlsx文件放到data目录下，命名为source.xlsx
2. 运行ExtractMovies.py
3. 运行CleanData.py
4. 打开CleanedMovies.xlsx，使用excel排序：主要B列（电影名），次要E列（导演），注意不要将标题栏一起排序
5. 运行DisposeDuplicates.py

### 第二阶段数据清洗操作
1. 将imdb爬取结果的xlsx文件放到data目录下，命名为ImdbTest.xlsx
2. 运行JoinImdb.py
3. 运行MergeAndClean.py
4. 运行SplitActors.py和SplitDirectors.py