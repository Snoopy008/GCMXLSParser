//
//  GCMViewController.m
//  GCMXLSParser
//
//  Created by 984603904@qq.com on 07/05/2022.
//  Copyright (c) 2022 984603904@qq.com. All rights reserved.
//

#import "GCMViewController.h"
#import <Masonry/Masonry.h>
#import <GCMXLSParser/GCMParserManager.h>

@interface GCMViewController ()

@end

@implementation GCMViewController

- (void)viewDidLoad
{
    [super viewDidLoad];
    self.view.backgroundColor = [UIColor whiteColor];
	// Do any additional setup after loading the view, typically from a nib.
    
    UIButton *csvBtn = [[UIButton alloc] init];
    csvBtn.backgroundColor = [UIColor blueColor];
    [csvBtn setTitle:@"CSV" forState:UIControlStateNormal];
    [csvBtn addTarget:self action:@selector(csvParser:) forControlEvents:UIControlEventTouchUpInside];
    [csvBtn setTitleColor:[UIColor whiteColor] forState:UIControlStateNormal];
    [self.view addSubview:csvBtn];
    [csvBtn mas_makeConstraints:^(MASConstraintMaker *make) {
        make.centerX.equalTo(self.view);
        make.top.equalTo(@80);
        make.size.mas_equalTo(CGSizeMake(100, 50));
    }];
    
    
    UIButton *xlsBtn = [[UIButton alloc] init];
    xlsBtn.backgroundColor = [UIColor blueColor];
    [xlsBtn setTitle:@"XLS" forState:UIControlStateNormal];
    [xlsBtn addTarget:self action:@selector(xlsParser:) forControlEvents:UIControlEventTouchUpInside];
    [xlsBtn setTitleColor:[UIColor whiteColor] forState:UIControlStateNormal];
    [self.view addSubview:xlsBtn];
    [xlsBtn mas_makeConstraints:^(MASConstraintMaker *make) {
        make.centerX.equalTo(self.view);
        make.top.equalTo(csvBtn.mas_bottom).offset(80);
        make.size.mas_equalTo(CGSizeMake(100, 50));
    }];
    
    
    UIButton *xlsxBtn = [[UIButton alloc] init];
    xlsxBtn.backgroundColor = [UIColor blueColor];
    [xlsxBtn setTitle:@"XLSX" forState:UIControlStateNormal];
    [xlsxBtn addTarget:self action:@selector(xlsxParser:) forControlEvents:UIControlEventTouchUpInside];
    [xlsxBtn setTitleColor:[UIColor whiteColor] forState:UIControlStateNormal];
    [self.view addSubview:xlsxBtn];
    [xlsxBtn mas_makeConstraints:^(MASConstraintMaker *make) {
        make.centerX.equalTo(self.view);
        make.top.equalTo(xlsBtn.mas_bottom).offset(80);
        make.size.mas_equalTo(CGSizeMake(100, 50));
    }];
    
    
    
    
}


- (void)csvParser:(UIButton *)sender
{
    NSString *path = [[NSBundle mainBundle] pathForResource:@"csvtest" ofType:@"csv"];
    
    [[GCMParserManager shareInstance] parserExcel_CSV_WithPath:path];
}


- (void)xlsParser:(UIButton *)sender
{
    NSString *path = [[NSBundle mainBundle] pathForResource:@"xlstest" ofType:@"xls"];
    
    [[GCMParserManager shareInstance] parserExcel_XLS_WithPath:path];
}

- (void)xlsxParser:(UIButton *)sender
{
    NSString *path = [[NSBundle mainBundle] pathForResource:@"xlsxtest" ofType:@"xlsx"];
    
    [[GCMParserManager shareInstance] parserExcel_XLSX_WithPath:path];
}




- (void)didReceiveMemoryWarning
{
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}

@end
