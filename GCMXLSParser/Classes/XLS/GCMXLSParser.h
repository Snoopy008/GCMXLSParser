//
//  GCMXLSParser.h
//  GCMXLSParser
//
//  Created by samuel on 2022/7/6.
//

#import <Foundation/Foundation.h>
#import "GCMCellModel.h"

NS_ASSUME_NONNULL_BEGIN

@interface GCMXLSParser : NSObject

@property (nonatomic, readonly) NSString *appName;
@property (nonatomic, readonly) NSString *author;
@property (nonatomic, readonly) NSString *category;
@property (nonatomic, readonly) NSString *comment;
@property (nonatomic, readonly) NSString *company;
@property (nonatomic, readonly) NSString *keywords;
@property (nonatomic, readonly) NSString *lastAuthor;
@property (nonatomic, readonly) NSString *manager;
@property (nonatomic, readonly) NSString *subject;
@property (nonatomic, readonly) NSString *title;

- (instancetype)initWithPath:(NSString *)filePath;

+ (NSString *)libaryVersion;

// Sheet Information
- (NSInteger)numberOfSheets;
- (NSString *)sheetNameAtIndex:(NSInteger)index;
- (BOOL)isSheetVisibleAtIndex:(NSUInteger)index;
- (NSInteger)numberOfRowsInSheet:(NSInteger)sheetIndex;
- (NSInteger)numberOfColsInSheet:(NSInteger)sheetIndex;

// Cell Accessor
// XLS row and column is 1 based, so use 1 based indexing here.
- (GCMCellModel *)cellInWorkSheetIndex:(NSInteger)sheetNum row:(NSInteger)row column:(NSInteger)column;
- (GCMCellModel *)cellInWorkSheetIndex:(NSInteger)sheetNum row:(NSInteger)row columnName:(NSString *)columnName; // "A"...."Z" "AA"..."ZZ"

// Iterate through all cells
- (void)startIteratorSheetAtIndex:(NSInteger)sheetIndex;
- (nullable GCMCellModel *)nextCell;

@end

NS_ASSUME_NONNULL_END
