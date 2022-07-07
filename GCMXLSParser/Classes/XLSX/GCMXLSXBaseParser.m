//
//  GCMXLSXBaseParser.m
//  GCMXLSParser
//
//  Created by samuel on 2022/7/7.
//

#import "GCMXLSXBaseParser.h"

@implementation GCMXLSXBaseParser

- (NSXMLParser *)xmlParser {
    
    if (!_xmlParser) {
        
        NSData *data = [[NSData alloc] initWithContentsOfFile:self.filePath];
        _xmlParser = [[NSXMLParser alloc] initWithData:data];
    }
    return _xmlParser;
}

@end
