//
//  GCMCellModel.m
//  GCMXLSParser
//
//  Created by samuel on 2022/7/5.
//

#import "GCMCellModel.h"

@implementation GCMCellModel

- (instancetype)init
{
    self = [super init];
    if (self) {
        self.number = @0;
    }
    return self;
}


- (NSString *)columnName
{
    if(_columnName == nil){
        char colStr[3];
        if(self.column < 26) {
            colStr[0] = 'A' + (char)self.column;
            colStr[1] = '\0';
        } else {
            colStr[0] = 'A' + (char)(self.column/26);
            colStr[1] = 'A' + (char)(self.column%26);
        }
        colStr[2] = '\0';
        NSString *columnName = [NSString stringWithFormat:@"%s", colStr];
        _columnName = columnName;

    }
    return _columnName;
    

}



- (NSString *)description {
    
    NSString *pointerInformation = [super description];
    NSDictionary *nameMapper = @{
                                 @(GCMCellContentTypeBlank): @"Blank",
                                 @(GCMCellContentTypeString): @"String",
                                 @(GCMCellContentTypeInteger): @"Integer",
                                 @(GCMCellContentTypeFloat): @"Float",
                                 @(GCMCellContentTypeBool): @"Bool",
                                 @(GCMCellContentTypeError): @"Error",
                                 };
    NSDictionary *dic = @{
                          @"type": nameMapper[@(self.contentType)],
                          @"row": @(self.row),
                          @"column": @(self.column),
                          @"columnName": self.columnName,
                          @"text": self.text,
                          @"value": self.number,
                          };
    NSString *desc = [NSString stringWithFormat:@"\n%@\n%@\n", pointerInformation, dic];
    return desc;
}



@end
