# GCMXLSParser

[![CI Status](https://img.shields.io/travis/984603904@qq.com/GCMXLSParser.svg?style=flat)](https://travis-ci.org/984603904@qq.com/GCMXLSParser)
[![Version](https://img.shields.io/cocoapods/v/GCMXLSParser.svg?style=flat)](https://cocoapods.org/pods/GCMXLSParser)
[![License](https://img.shields.io/cocoapods/l/GCMXLSParser.svg?style=flat)](https://cocoapods.org/pods/GCMXLSParser)
[![Platform](https://img.shields.io/cocoapods/p/GCMXLSParser.svg?style=flat)](https://cocoapods.org/pods/GCMXLSParser)

## Example

To run the example project, clone the repo, and run `pod install` from the Example directory first.

## Requirements

## Installation

GCMXLSParser is available through [CocoaPods](https://cocoapods.org). To install
it, simply add the following line to your Podfile:

```ruby
pod 'GCMXLSParser'
```


## 使用
```
#import <GCMXLSParser/GCMParserManager.h>
...
//CSV文件解析
NSArray *array =[[GCMParserManager shareInstance] parserExcel_CSV_WithPath:path];

//XLS文件解析
NSArray *array =[[GCMParserManager shareInstance] parserExcel_XLS_WithPath:path];

//XLSX文件解析
NSArray *array =[[GCMParserManager shareInstance] parserExcel_XLSX_WithPath:path];
```
如果觉得比较耗时可以将解析方法放在异步函数内
如果有问题可以提issue,好用记得点赞和打赏哦！
您的鼓励和支持是我最大的动力。

## Author

984603904@qq.com,

## License

GCMXLSParser is available under the MIT license. See the LICENSE file for more info.
