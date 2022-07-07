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
NSString *path = [[NSBundle mainBundle] pathForResource:@"test3" ofType:@"xlsx"];
[LAWExcelTool shareInstance].delegate = self;
[[LAWExcelTool shareInstance] parserExcelWithPath:path];
```

## Author

984603904@qq.com,

## License

GCMXLSParser is available under the MIT license. See the LICENSE file for more info.
