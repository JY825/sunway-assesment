import {
  Component,
  ElementRef,
  Input,
  Output,
  EventEmitter,
  ViewChild,
} from "@angular/core";
import { Router } from "@angular/router";
import { GlobalDataStorage } from "src/app/share/storages/global-data.storage";
import { FundService } from "src/app/share/services/fund.service";
import { PublicHolidayService } from "src/app/share/services/publicholiday.service";
import { DateTimeConstant } from "src/app/share/constants/date-time.constant";
import { TranslateService } from "@ngx-translate/core";
import { forkJoin } from "rxjs";
import {Chart} from 'angular-highcharts';
import { ActivatedRoute } from "@angular/router";
import { NzMessageService } from 'ng-zorro-antd/message';
import { Location } from "@angular/common";
import { dateFormat, getOptions, setOptions } from 'highcharts';
import * as Moment from "moment";
import { ExcelService } from "../../../share/services/excel.service";
import { PdfService } from "../../../share/services/pdf.service";
import { WatchlistService } from 'src/app/share/services/watchlist.service';
import { UiModalService } from '@share/ui-components/modal/ui-modal.service';
import * as pdfMake from "pdfmake/build/pdfmake";
import * as FileSaver from "file-saver-es";
import * as html2pdf from "html2pdf.js";
import { ResponsiveService } from "src/app/share/services/responsive.service";
import {DateFormatPipe} from "../../../share/pipes/date-format.pipe";
import { RcsService } from "src/app/share/services/rcs.service";
import { UtilService } from "@share/services/util.service";
import { FundChartService } from "@share/services/fund-chart.service";

window["moment"] = Moment;

const EXCEL_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8";
const EXCEL_EXTENSION = ".xlsx";

setOptions({
  time: {
    useUTC: false,
  },
});

interface TopHoldingsItem {
  name: string;
  percentage: number;
}

interface ITabModel {
  chartPeriod: string;
  value: string;
  label: string;
  active?: boolean;
}

interface PerformancePoint {
  date: Date;
  value: number;
}

type PeriodType = '1M' | '3M' | '6M' | '1Y' | '3Y' | '5Y' | '10Y';

@Component({
  templateUrl: "./fund-factsheet.component.html",
  styleUrls: ["./fund-factsheet.component.scss"],
})
export class FundFactsheetComponent {
  @ViewChild("resultContainer") resultContainer: ElementRef;
  @ViewChild("resultContainer2") resultContainer2: ElementRef;
  pageName = "FundFactSheet";
  @Input() hideExportExcel: boolean;
  @Input() refno: string;
  @Output() onClickPdfAndPrint: EventEmitter<any> = new EventEmitter();

  isLoadingPriceAlert: boolean = false;
  showPriceAlert: boolean = false;
  ceilingPrice: any = 0;
  floorPrice: any = 0;

  isDownloadLoading: boolean = false;
  listAllClient: boolean = false;
  faAccData: any;
  watchListObj = {};
  riskImageLink: string = "/b2b/fsm/images/en/riskRatingOverview.jpg";

  fundDetailsLoading: boolean = false;
  fundPriceInfoLoading: boolean = false;
  fundSectorPerformanceLoading: boolean = false;
  historicSharpeLoading: boolean = false;
  fundConsistLoading: boolean = false;
  priceHistoryLoading: boolean = false;
  fundCalendarLoading: boolean = false;
  publicHolidayLoading: boolean = false;
  fundReturnsLoading: boolean = false;
  performanceChartLoading: boolean = false;
  historicalDividendLoading: boolean = false;
  fundFilesLoading: boolean = true;

  isDailyPriceHistoryModalVisible: boolean = false;
  isRiskRatingModalVisible: boolean = false;
  isHistoricalDividendVisible: boolean = false;
  isComplexProduct: boolean = false;
  isCFAAuthorised: boolean = false;
  fundDetails: any;
  fundPriceInfo: any;
  dailyPriceHistory: any;
  displayDailyPriceHistoryData: any;
  displayHistoricalData: any;

  dailyPriceDataSortMap = {
    bidPriceTest: null,
    showDateTest: null,
  };
  historicalDividendDataSortMap = {
    exDateStr: null,
    rateStr: null,
    divYield: null,
  };

  sectorCurrencyType: string = "baseCurrency";
  performanceCurrencyType: string = "baseCurrency";
  fundSectorPerformance: any;
  fundSectorPerformanceFc: any;
  historicSharpe: any;
  historicSharpeFc: any;
  fundConsist: any;
  fundConsistFc: any;
  selectedFundSectorPerformance: any;
  selectedFundSectorPerformance2: any;
  selectedHistoricSharpe: any;
  selectedFundConsist: any;
  fundCalendarList: any;
  publicHolidayList: any;
  fundReturnsResult: any;
  fundReturnsErr: any;
  resultCurrencyType: string = "baseCurrency";
  performanceChart: any;
  performanceChartFc: any;
  selectedPerformanceChart: any;
  selectFundReturnsPerformance: any;
  fundConsistChart: any;
  fundManager: any;
  dividendHistoryList: any;
  fundId: string;

  errorMsg: string;
  dateTimeFormat: string;

  dateRangeFromDate: any = null;
  dateRangeToDate: any = null;

  baseCurrency: string = "hkd";
  performanceSelectedCurrency: string = "hkd";

  searchDailyPriceHistoryText: string;
  searchHistoricalDividendText: string;

  dailyPriceHistoryPageSize = 10;
  historicalDividendPageSize = 10;

  emptyDataMessage = " ";

  dateFormat = "yyyy/MM/dd";

  selectedLocale: string;
  displayFactsheet: boolean = true;

  exportLoading: boolean = false;

  prospectusPath: String = "";
  reportPath: string = "";
  kfsPath: string = "";
  factsheetPath: string = "";

  isShareTooltip: boolean = false;

  displayTransferInOnlyFund: boolean = false;

  radioButtonStyle = {
    display: "block",
    height: "30px",
    lineHeight: "30px",
  };

  performanceChartConfig: any;
  sectorPerformanceChart = new Chart({
    chart: {
      type: "column",
      plotBackgroundColor: null,
      zooming: {
        type: 'x'
      },
      backgroundColor: null,
      borderWidth: 0,
      style: {
        overflow: "visible",
      },
    },
    title: {
      text: "",
    },
    yAxis: {
      endOnTick: false,
      startOnTick: false,
      labels: {
        enabled: false,
      },
      title: {
        text: null,
      },
      tickPositions: [0],
    },
    xAxis: {
      title: {
        text: null,
      },
      labels: {
        enabled: false,
      },
      startOnTick: false,
      endOnTick: false,
      lineWidth: 0,
      minorGridLineWidth: 0,
      tickPositions: [],
    },
    credits: {
      enabled: false,
    },
    legend: {
      enabled: false,
    },
    tooltip: {
      formatter: function () {
        return "<b>" + this.key + ": </b>" + this.y.toFixed(2) + "%";
      },
    },
    plotOptions: {
      series: {
        animation: false,
        lineWidth: 1,
        shadow: false,
        states: {
          hover: {
            lineWidth: 1,
          },
        },
        marker: {
          radius: 1,
          states: {
            hover: {
              radius: 2,
            },
          },
        },
      },
      column: {
        negativeColor: "#910000",
      },
    },
  });

  errorCodeMap = {
    ERROR_FUND_NOT_FOUND: "error.fund.not.found",
  };

  chartPeriodOptions: ITabModel[] = [];

  isMobileView: boolean;
  mobileViewSubscription: any;

  displayRegionalMargin = this.globalDataStorage.displayRegionalMargin;
  productLtvRatio: number;
  productLtvRatioString: string;
  marginable: boolean;
  marginTradable: boolean;

  fundTopHoldingsList: TopHoldingsItem[] = [];
  fundTopHoldingsLastUpdateDate: any;
  readonly allPeriodTabs: PeriodType[] = ['1M', '3M', '6M', '1Y', '3Y', '5Y', '10Y'];
  availablePeriodTabs: PeriodType[] = [];
  performanceChartData: PerformancePoint[] = [];
  performanceSelectedPeriod: PeriodType = '1Y';
  selectedChartPeriod: string = '3y';
  chartPeriodLoading: boolean = false;

  constructor(
    private router: Router,
    private fundService: FundService,
    private publicHolidayService: PublicHolidayService,
    private globalDataStorage: GlobalDataStorage,
    private translateService: TranslateService,
    private route: ActivatedRoute,
    private datePipe: DateFormatPipe,
    private messageService: NzMessageService,
    private location: Location,
    private excelService: ExcelService,
    private pdfService: PdfService,
    private watchlistService: WatchlistService,
    private nzModalService: UiModalService,
    private responsiveService: ResponsiveService,
    private rcsService: RcsService,
    private utilService: UtilService,
    private fundChartService:FundChartService,
  ) {
    this.dateTimeFormat = DateTimeConstant.dateTimeFormat;
    this.isMobileView = this.responsiveService.isMobileView();
    this.mobileViewSubscription = this.responsiveService.mobileViewEvent.subscribe((state) => {
      this.isMobileView = state;
    });
  }

  ngOnInit(): void {
    this.fundId = this.route.snapshot.paramMap.get("fundId");
    this.globalDataStorage.selectedFund = null;
    this.selectedLocale = this.globalDataStorage.getSessionStorage("locale") || "en";
    this.initializeChartPeriodOptions();
    this.getFundBasicInfo(this.fundId);
    this.getFundSectorPerformance(this.fundId);
    this.getHistoricSharpe(this.fundId);
    this.getFundConsist(this.fundId);
    this.getDailyPriceHistory(this.fundId);
    this.getNext10FundHoliday(this.fundId);
    this.getNext10PublicHoliday();
    this.loadInitialPerformanceCharts();
    this.getFundManager(this.fundId);
    this.findProductLtvInfo(this.fundId);
    this.loadChartDataByPeriod(this.selectedChartPeriod);
  }

  ngOnDestroy(): void {
    if (this.mobileViewSubscription) {
      this.mobileViewSubscription.unsubscribe();
    }
  }

  shareURL() {
    window.navigator['clipboard'].writeText(
      "https://www.ifastfinancial.com.hk/b2b-adviser/post-login/fund-factsheet" +
        this.fundId
    );
    this.isShareTooltip = true;
    if(this.isShareTooltip){
      setTimeout(() => (this.isShareTooltip = false) , 2000)
    }
  }

  initializeChartPeriodOptions(): void {
    this.chartPeriodOptions = [
      {
        chartPeriod: '1m',
        value: '1m',
        label: this.translateService.instant('chart.one.month')
      },
      {
        chartPeriod: '3m',
        value: '3m',
        label: this.translateService.instant('chart.three.months')
      },
      {
        chartPeriod: '6m',
        value: '6m',
        label: this.translateService.instant('chart.six.months')
      },
      {
        chartPeriod: '1y',
        value: '1y',
        label: this.translateService.instant('chart.one.year')
      },
      {
        chartPeriod: '3y',
        value: '3y',
        label: this.translateService.instant('chart.three.years'),
        active: true
      },
      {
        chartPeriod: '5y',
        value: '5y',
        label: this.translateService.instant('chart.five.years')
      },
      {
        chartPeriod: '10y',
        value: '10y',
        label: this.translateService.instant('chart.ten.years')
      }
    ];
  }

  loadInitialPerformanceCharts(): void {
    this.performanceChartLoading = true;
    
    forkJoin([
      this.fundService.getPerformanceChart(this.fundId, "N"),
      this.fundService.getPerformanceChart(this.fundId, "Y")
    ]).subscribe(
      ([performanceChart, performanceChartFc]) => {
        if (performanceChart.success) {
          this.performanceChart = performanceChart.data;
          this.selectedPerformanceChart = this.performanceChart;
          
          if (this.fundDetails != null) {
            this.updatePerfChartData(this.selectedPerformanceChart);
          }
        }

        if (performanceChartFc.success) {
          this.performanceChartFc = performanceChartFc.data;
        }
        
        this.performanceChartLoading = false;
      },
      (error) => {
        this.errorMsg = "Error occurs";
        this.performanceChartLoading = false;
      }
    );
  }

  // CORRECTED METHOD
  onChangeChartPeriod(selectedPeriod: string): void {
    if (this.selectedChartPeriod === selectedPeriod || this.chartPeriodLoading) {
      return;
    }

    this.selectedChartPeriod = selectedPeriod;
    this.chartPeriodOptions.forEach(option => {
      option.active = option.chartPeriod === selectedPeriod;
    });

    // Load data from backend for that period
    this.loadChartDataByPeriod(selectedPeriod);
  }

  // CORRECTED METHOD
  loadChartDataByPeriod(period: string): void {
    this.chartPeriodLoading = true;

    // Call the new API instead of filtering old data
    this.fundChartService.getFundChartCentreData(this.fundId, period).subscribe({
      next: (res: any) => {
        if (res?.success && Array.isArray(res.data)) {
          const rawData: PerformancePoint[] = res.data.map((point: [number, number]) => ({
            date: new Date(point[0]),
            value: point[1]
          }));

          this.performanceChartData = rawData.sort(
            (a, b) => a.date.getTime() - b.date.getTime()
          );

          this.updateChartWithApiData();
        } else {
          // If API fails, try to use existing chart data with filtering
          if (this.performanceChart && this.performanceChart.data) {
            const filteredData = this.filterChartDataByPeriod(this.performanceChart.data, period);
            this.updatePerfChartDataWithFiltered(filteredData);
          } else if (this.selectedPerformanceChart) {
            const filteredData = this.filterChartDataByPeriod(this.selectedPerformanceChart, period);
            this.updatePerfChartDataWithFiltered(filteredData);
          }
        }

        this.chartPeriodLoading = false;
      },
      error: (error) => {
        console.error('Error loading chart data:', error);
        
        // Fallback to existing data filtering if API fails
        if (this.performanceChart && this.performanceChart.data) {
          const filteredData = this.filterChartDataByPeriod(this.performanceChart.data, period);
          this.updatePerfChartDataWithFiltered(filteredData);
        } else if (this.selectedPerformanceChart) {
          const filteredData = this.filterChartDataByPeriod(this.selectedPerformanceChart, period);
          this.updatePerfChartDataWithFiltered(filteredData);
        }
        
        this.chartPeriodLoading = false;
      }
    });
  }

  // NEW METHOD
  updateChartWithApiData(): void {
    if (!this.performanceChartData || this.performanceChartData.length === 0) {
      return;
    }

    const seriesData = this.performanceChartData.map(d => [d.date.getTime(), d.value]);
    const fundName = this.fundDetails?.fundName || 'Fund Performance';

    this.performanceChartConfig = new Chart({
      chart: {
        plotBorderWidth: null,
        plotShadow: false,
        type: "line",
        height: 350,
        zooming: {
          type: 'x'
        }
      },
      title: {
        text: "",
      },
      yAxis: {
        title: {
          text: this.translateService.instant('return.percentage'),
        },
        labels: {
          format: '{value}%'
        }
      },
      xAxis: {
        type: "datetime",
        crosshair: true
      },
      credits: {
        enabled: false,
      },
      legend: {
        enabled: false,
      },
      tooltip: {
        formatter: function () {
          return (
            dateFormat("%d %b %Y", this.x as number) +
            '<br/><span style="color: ' + '">' +
            fundName +
            ": </span><b>" +
            this.y + "%" +
            "</b>"
          );
        },
      },
      plotOptions: {
        area: {
          fillColor: {
            linearGradient: { x1: 0, y1: 0, x2: 0, y2: 1 },
            stops: [
              [0, getOptions().colors[0] as string],
              [1, "#ffffff"],
            ],
          },
        },
        series: {
          marker: {
            enabled: false,
          },
        },
      },
      series: [
        {
          type: 'line',
          name: fundName,
          data: seriesData,
        },
      ],
    });
  }

  // NEW METHOD
  updatePerfChartDataWithFiltered(chartData: any): void {
    if (!chartData || !this.fundDetails) {
      return;
    }
    
    const fundName = this.fundDetails.fundName;
    this.performanceChartConfig = new Chart({
      chart: {
        plotBorderWidth: null,
        plotShadow: false,
        type: "line",
        height: 350,
        zooming: {
          type: 'x'
        }
      },
      title: {
        text: "",
      },
      yAxis: {
        title: {text: this.translateService.instant('return.percentage'),},
        labels: {
          format: '{value}%'
        }
      },
      xAxis: {
        type: "datetime",
        crosshair: true
      },
      credits: {
        href: "/fsm/main/home",
        text: "",
      },
      legend: {
        enabled: false,
      },
      tooltip: {
        formatter: function () {
          return (
            dateFormat("%d %b %Y", this.x as number) +
            '<br/><span style="color: ' + '">' +
            fundName +
            ": </span><b>" +
            this.y + "%" +
            "</b>"
          );
        },
      },
      plotOptions: {
        area: {
          fillColor: {
            linearGradient: { x1: 0, y1: 0, x2: 0, y2: 1 },
            stops: [
              [0, getOptions().colors[0] as string],
              [1, "#ffffff"],
            ],
          },
        },
        series: {
          marker: {
            enabled: false,
          },
        },
      },
      series: [
        {
          type: 'line',
          name: this.fundDetails.fundName,
          data: chartData,
        },
      ],
    });
  }

  filterChartDataByPeriod(chartData: any[], period: string): any[] {
    if (!chartData || !Array.isArray(chartData)) {
      return chartData;
    }

    const currentDate = new Date();
    let cutoffDate = new Date();

    switch (period) {
      case '1m':
        cutoffDate.setMonth(currentDate.getMonth() - 1);
        break;
      case '3m':
        cutoffDate.setMonth(currentDate.getMonth() - 3);
        break;
      case '6m':
        cutoffDate.setMonth(currentDate.getMonth() - 6);
        break;
      case '1y':
        cutoffDate.setFullYear(currentDate.getFullYear() - 1);
        break;
      case '3y':
        cutoffDate.setFullYear(currentDate.getFullYear() - 3);
        break;
      case '5y':
        cutoffDate.setFullYear(currentDate.getFullYear() - 5);
        break;
      case '10y':
        cutoffDate.setFullYear(currentDate.getFullYear() - 10);
        break;
      default:
        return chartData;
    }

    return chartData.filter(dataPoint => {
      if (Array.isArray(dataPoint) && dataPoint.length >= 2) {
        const dataDate = new Date(dataPoint[0]);
        return dataDate >= cutoffDate;
      }
      return true;
    });
  }

  // Rest of your methods remain the same...
  findProductLtvInfo(fundId) {
    this.rcsService.findProductLtvInfo(fundId, "UT").subscribe(result => {
      if (result.success) {
        this.productLtvRatio = result.data[0].ltvRatio;
        this.productLtvRatioString = result.data[0].ltvRatioString;
        this.marginable = result.data[0].marginable === 'Y';
        this.marginTradable = result.data[0].marginTradable === 'Y';
      }
    })
  }

  goToTransactPage(fundId) {
    this.globalDataStorage.selectedFund = fundId;
    const url = this.router.serializeUrl(this.router.createUrlTree(["/b2b-adviser/post-login/trades/select-client-account"], {queryParams: { link: "fund-subscription" }}));
    window.open(url, '_blank');
  }

  goToRSPApplication(fundId) {
    this.globalDataStorage.selectedFund = fundId;
    const url = this.router.serializeUrl(this.router.createUrlTree(["/b2b-adviser/post-login/trades/select-client-account"], {queryParams: { link: "fund-rsp"}}))
    window.open(url, '_blank');
  }

  decideDisplay(fundDetails) {
    this.displayFactsheet = false;
    if (null != this.globalDataStorage.platform) {
      if (
        ("igp" == this.globalDataStorage.platform &&
          "Y" == fundDetails.distributeIgp) ||
        ("ifa" == this.globalDataStorage.platform &&
          "Y" == fundDetails.distributeIfast) ||
        ("ins" == this.globalDataStorage.platform &&
          "Y" == fundDetails.distributeInsti) ||
        this.displayTransferInOnlyFund == true
      ) {
        this.displayFactsheet = true;
      }
    }
  }

  getFundBasicInfo(fundId: string) {
    this.fundDetailsLoading = true;
    this.fundPriceInfoLoading = true;
    forkJoin([
      this.fundService.getFundDetails(fundId),
      this.fundService.getFundPriceInfo(fundId),
      this.fundService.getDividendHistoryList(fundId),
    ]).subscribe(
      ([fundDetails, fundPriceInfo, dividendHistory]) => {
        if (fundDetails.success) {
          this.fundDetails = fundDetails.data;
          this.fundTopHoldingsList = this.fundDetails.topHoldingList;
          this.fundTopHoldingsLastUpdateDate = this.fundDetails.topHoldingDate;
          this.fundDetailsLoading = false;
          if (this.fundDetails.transferInOnly == "Y") {
            this.displayTransferInOnlyFund = true;
          }
          if (this.fundDetails.complexProduct == "Y") {
            this.isComplexProduct = true;
          }
          if (
            this.fundDetails.classificationScheme.toLowerCase() ==
            "SFC-Authorised".toLowerCase()
          ) {
            this.isCFAAuthorised = true;
          }
          if (this.selectedPerformanceChart != null) {
            this.updatePerfChartData(this.selectedPerformanceChart);
          }
          
          this.decideDisplay(this.fundDetails);
          this.getFundFilesExists(this.fundId);
        } else {
          this.errorMsg = fundDetails.message;
          this.messageService.create(
            "error",
            this.translateService.instant(
              "" + this.errorCodeMap["" + this.errorMsg]
            )
          );
          this.location.back();
        }

        if (fundPriceInfo.success) {
          this.fundPriceInfo = fundPriceInfo.data;
          this.ceilingPrice = this.fundPriceInfo.dailyPrice;
          this.floorPrice = this.fundPriceInfo.dailyPrice;
          this.fundPriceInfoLoading = false;
        } else {
          this.errorMsg = fundPriceInfo.message;
          this.messageService.create(
            "error",
            this.translateService.instant(
              "" + this.errorCodeMap["" + this.errorMsg]
            )
          );
          this.location.back();
        }

        if (dividendHistory.success) {
          this.dividendHistoryList = dividendHistory.data;
          this.dividendHistoryList.forEach((dh) => {
            dh.exDateStr = "" + dh.exDate;
            dh.currency = "" + dh.currency;
            dh.rateStr = "" + dh.rate;
            dh.divYield = "" + dh.divYield;
          });
          this.displayHistoricalData = this.dividendHistoryList;
          this.setExportContent();
        }

        this.fundDetailsLoading = false;
      },
      (error) => {
        this.errorMsg = "Error occurs";
      }
    );
  }

  getFundSectorPerformance(fundId: string) {
    this.fundSectorPerformanceLoading = true;
    forkJoin([
      this.fundService.getFundSectorPerformance(fundId, "N"),
      this.fundService.getFundSectorPerformance(fundId, "Y"),
    ]).subscribe(
      ([fundSectorPerformance, fundSectorPerformanceFc]) => {
        if (fundSectorPerformance.success) {
          this.fundSectorPerformance = fundSectorPerformance.data;
          var series = [];
          this.fundSectorPerformance.chartData.forEach(
            (subData: any, subKey: any) => {
              var value = [];
              value.push(subData[0]);
              value.push(subData[1]);
              series.push(value);
            }
          );
          this.fundSectorPerformance.chartDataTest = series;
          this.selectedFundSectorPerformance = this.fundSectorPerformance;
          this.selectedFundSectorPerformance2 = this.fundSectorPerformance;
          this.updateChartData(
            this.selectedFundSectorPerformance.chartDataTest
          );
        } else {
          this.errorMsg = fundSectorPerformance.message;
          this.messageService.create(
            "error",
            this.translateService.instant(
              "" + this.errorCodeMap["" + this.errorMsg]
            )
          );
          this.location.back();
        }

        if (fundSectorPerformanceFc.success) {
          this.fundSectorPerformanceFc = fundSectorPerformanceFc.data;
          var series = [];
          this.fundSectorPerformanceFc.chartData.forEach(
            (subData: any, subKey: any) => {
              var value = [];
              value.push(subData[0]);
              value.push(subData[1]);
              series.push(value);
            }
          );
          this.fundSectorPerformanceFc.chartDataTest = series;
        }
        this.fundSectorPerformanceLoading = false;
      },
      (error) => {
        this.errorMsg = "Error occurs";
      }
    );
  }

  getHistoricSharpe(fundId: string) {
    this.historicSharpeLoading = true;
    forkJoin([
      this.fundService.getHistoricSharpe(fundId, "N"),
      this.fundService.getHistoricSharpe(fundId, "Y"),
    ]).subscribe(
      ([historcSharpe, historicSharpeFc]) => {
        if (historcSharpe.success) {
          this.historicSharpe = historcSharpe.data;
          this.selectedHistoricSharpe = this.historicSharpe;
        }

        if (historicSharpeFc.success) {
          this.historicSharpeFc = historicSharpeFc.data;
        }
        this.historicSharpeLoading = false;
      },
      (error) => {
        this.errorMsg = "Error occurs";
      }
    );
  }

  getFundConsist(fundId: string) {
    this.fundConsistLoading = true;
    forkJoin([
      this.fundService.getFundConsist(fundId, "N"),
      this.fundService.getFundConsist(fundId, "Y"),
    ]).subscribe(
      ([fundConsist, fundConsistFc]) => {
        if (fundConsist.success) {
          this.fundConsist = fundConsist.data;
          this.selectedFundConsist = this.fundConsist;
          this.updateFundConsistChartData(this.selectedFundConsist.chartData);
        }

        if (fundConsistFc.success) {
          this.fundConsistFc = fundConsistFc.data;
        }
        this.fundConsistLoading = false;
      },
      (error) => {
        this.errorMsg = "Error occurs";
      }
    );
  }

  getDailyPriceHistory(fundId: string) {
    this.priceHistoryLoading = true;
    this.fundService.getDailyPriceHistory(fundId).subscribe((result) => {
      if (result.success) {
        this.dailyPriceHistory = result.data;
        this.dailyPriceHistory.forEach((dph) => {
          dph.bidPriceTest = "" + dph.bidPrice;
          dph.offerPriceTest = "" + dph.offerPrice;
          dph.showDateTest =
            "" +
              this.datePipe.transform(dph.dailyPricePk.showDate, "mediumDate");
        });
        this.displayDailyPriceHistoryData = this.dailyPriceHistory;
      }
      this.priceHistoryLoading = false;
      this.setExportContent2();
    });
  }

  getNext10FundHoliday(fundId: string) {
    this.fundCalendarLoading = true;
    this.fundService.getFundHoliday(fundId, 10).subscribe((result) => {
      if (result.success) {
        this.fundCalendarList = result.data;
      }
      this.fundCalendarLoading = false;
    });
  }

  getNext10PublicHoliday() {
    this.publicHolidayLoading = true;
    this.publicHolidayService
      .findNextDaysPublicHolidays(10)
      .subscribe((result) => {
        if (result.success && result.length > 0 ) {
          this.publicHolidayList = result.data;
        }else if (result.length == 0) {
          this.publicHolidayList = false;
        }
        this.publicHolidayLoading = false;
      });
  }

  getFundReturns(fundId: string, fromDate: string, toDate: string) {
    this.fundReturnsLoading = true;
    this.fundService
      .getFundReturns(fundId, fromDate, toDate)
      .subscribe((result) => {
        if (result.success) {
          this.fundReturnsResult = result.data;
          if (this.fundReturnsResult != null) {
            this.selectFundReturnsPerformance =
              this.fundReturnsResult.performanceLocal;
          } else {
            this.fundReturnsErr = "";
          }
        }
        this.fundReturnsLoading = false;
      });
  }

  getPerformanceChart(fundId: string, isForeignCurrency: string) {
    // Use the new API
    this.fundChartService.getFundChartCentreData(fundId, isForeignCurrency).subscribe((res: any) => {
      if (res?.success && Array.isArray(res.data)) {
        // Assuming the new API returns [timestamp, value] pairs just like before
        const rawData: PerformancePoint[] = res.data.map((point: [number, number]) => ({
          date: new Date(point[0]),
          value: point[1]
        }));

        // Sort ascending by date
        this.performanceChartData = rawData.sort((a, b) => a.date.getTime() - b.date.getTime());

        // Recalculate periods and update chart for selected period
        this.calculateAvailablePeriods();
        this.updateChartForPeriod(this.performanceSelectedPeriod);
      } else {
        // Reset data if API fails or no data
        this.performanceChartData = [];
        this.performanceChart = null;
        this.availablePeriodTabs = [];
      }
    });
  }

  calculateAvailablePeriods() {
    const firstDate = this.performanceChartData[0]?.date;
    const lastDate = this.performanceChartData[this.performanceChartData.length - 1]?.date;

    if (!firstDate || !lastDate) {
      this.availablePeriodTabs = [];
      return;
    }

    const monthDiff = (lastDate.getFullYear() - firstDate.getFullYear()) * 12 + (lastDate.getMonth() - firstDate.getMonth());

    this.availablePeriodTabs = this.allPeriodTabs.filter(period => {
      switch (period) {
        case '1M': return monthDiff >= 1;
        case '3M': return monthDiff >= 3;
        case '6M': return monthDiff >= 6;
        case '1Y': return monthDiff >= 12;
        case '3Y': return monthDiff >= 36;
        case '5Y': return monthDiff >= 60;
        case '10Y': return monthDiff >= 120;
        default: return false;
      }
    });

    if (!this.availablePeriodTabs.includes(this.performanceSelectedPeriod)) {
      this.performanceSelectedPeriod = this.availablePeriodTabs[0];
    }
  }

  updateChartForPeriod(period: PeriodType) {
    const cutoffDate = new Date();
    switch (period) {
      case '1M': cutoffDate.setMonth(cutoffDate.getMonth() - 1); break;
      case '3M': cutoffDate.setMonth(cutoffDate.getMonth() - 3); break;
      case '6M': cutoffDate.setMonth(cutoffDate.getMonth() - 6); break;
      case '1Y': cutoffDate.setFullYear(cutoffDate.getFullYear() - 1); break;
      case '3Y': cutoffDate.setFullYear(cutoffDate.getFullYear() - 3); break;
      case '5Y': cutoffDate.setFullYear(cutoffDate.getFullYear() - 5); break;
      case '10Y': cutoffDate.setFullYear(cutoffDate.getFullYear() - 10); break;
    }

    const filteredData = this.performanceChartData.filter(d => d.date >= cutoffDate);
    const seriesData = filteredData.map(d => [d.date.getTime(), d.value]);

    // Use new Chart() for angular-highcharts
    this.performanceChartConfig = new Chart({
      chart: {
        type: 'line'
      },
      title: {
        text: ''
      },
      xAxis: {
        type: 'datetime',
        title: { text: 'Date' }
      },
      yAxis: {
        title: { text: 'Value' }
      },
      series: [{
        type: 'line', // <-- THIS IS IMPORTANT
        name: 'Performance',
        data: seriesData
      }]
    });
  }

  onPeriodTabChange(period: PeriodType) {
    this.performanceSelectedPeriod = period;
    this.updateChartForPeriod(period);
  }

  onChangeDateRangeFromDate(result: Date) {
    this.dateRangeFromDate = result;
  }

  onChangeDateRangeToDate(result: Date) {
    this.dateRangeToDate = result;
  }

  generateFundReturns() {
    var dateRangeFromVal = this.dateRangeFromDate.getTime();
    var dateRangeToVal = this.dateRangeToDate.getTime();

    this.fundReturnsErr = null;
    this.selectFundReturnsPerformance = null;

    if (dateRangeFromVal > dateRangeToVal) {
      this.fundReturnsErr =
        "fund.factsheet.from.date.should.be.smaller.than.the.to.date";
    } else {
      var one_day = 1000 * 60 * 60 * 24;
      var dateDifference = Math.ceil(
        (dateRangeToVal - dateRangeFromVal) / one_day
      );
      if (dateDifference < 31) {
        this.fundReturnsErr = "error.of.short.day.range";
      } else {
        this.getFundReturns(this.fundId, dateRangeFromVal, dateRangeToVal);
      }
    }
  }

  getFundManager(fundId: string) {
    this.fundReturnsLoading = true;
    this.fundService.getFundManager(fundId).subscribe((result) => {
      if (result.success) {
        this.fundManager = result.data;
        if (this.fundManager.url != null) {
          this.fundManager.url = this.fundManager.url.replace("https://", "");
          this.fundManager.url = this.fundManager.url.replace("http://", "");
        }
      }
    });
  }

  getFundFilesExists(fundId: string) {
    this.fundService.getFundFilesExists(fundId).subscribe((result) => {
      if (result.success) {
        this.selectedLocale =
          this.globalDataStorage.getSessionStorage("locale") || "en";
        if (result.data.prospectusDocumentExists) {
          this.fundDetails.prospectusDocumentExists =
            result.data.prospectusDocumentExists;
          this.prospectusPath = result.data.prospectusPath;
        }

        if (result.data.factsheetDocumentExists) {
          this.fundDetails.factsheetDocumentExists =
            result.data.factsheetDocumentExists;
          this.factsheetPath = result.data.factsheetPath;
        }

        if (result.data.fundReportsDocumentExists) {
          this.fundDetails.fundReportsDocumentExists =
            result.data.fundReportsDocumentExists;
          this.reportPath = result.data.fundReportsPath;
        }

        if (result.data.fundKfsDocumentExists) {
          this.fundDetails.fundKfsDocumentExists =
            result.data.fundKfsDocumentExists;
          this.kfsPath = result.data.fundKfsPath;
        }
      }
      this.fundFilesLoading = false;
    });
  }

  changeSectorPerformance(selectedSectorCurrency: string) {
    if (selectedSectorCurrency === "baseCurrency") {
      this.sectorCurrencyType = selectedSectorCurrency;
      this.selectedFundSectorPerformance = this.fundSectorPerformance;
    } else if (selectedSectorCurrency === "foreignCurrency") {
      this.sectorCurrencyType = selectedSectorCurrency;
      this.selectedFundSectorPerformance = this.fundSectorPerformanceFc;
    }
    this.updateChartData(this.selectedFundSectorPerformance.chartDataTest);
  }

  changePerformance(selectedPerformanceCurrency: string) {
    if (selectedPerformanceCurrency === "baseCurrency") {
      this.performanceCurrencyType = selectedPerformanceCurrency;
      this.performanceSelectedCurrency = "HKD";
      this.selectedFundSectorPerformance2 = this.fundSectorPerformance;
      this.selectedFundConsist = this.fundConsist;
      this.selectedHistoricSharpe = this.historicSharpe;
      
      if (this.performanceChart) {
        this.selectedPerformanceChart = this.filterChartDataByPeriod(this.performanceChart, this.selectedChartPeriod);
      }
    } else if (selectedPerformanceCurrency === "foreignCurrency") {
      this.performanceCurrencyType = selectedPerformanceCurrency;
      this.performanceSelectedCurrency = this.fundDetails.fundCurrencyCode;
      this.selectedFundSectorPerformance2 = this.fundSectorPerformanceFc;
      this.selectedFundConsist = this.fundConsistFc;
      this.selectedHistoricSharpe = this.historicSharpeFc;
      
      if (this.performanceChartFc) {
        this.selectedPerformanceChart = this.filterChartDataByPeriod(this.performanceChartFc, this.selectedChartPeriod);
      }
    }
    
    this.updateFundConsistChartData(this.selectedFundConsist.chartData);
    if (this.selectedPerformanceChart) {
      this.updatePerfChartData(this.selectedPerformanceChart);
    }
  }

  updateChartData(chartData: any) {
    this.sectorPerformanceChart = new Chart({
      chart: {
        type: "column",
        plotBackgroundColor: null,
        zooming: {
          type: 'x'
        },
        backgroundColor: null,
        borderWidth: 0,
        width: 120,
        height: 160,
        style: {
          overflow: "visible",
        },
      },
      title: {
        text: "",
      },
      yAxis: {
        endOnTick: false,
        startOnTick: false,
        labels: {
          enabled: false,
        },
        title: {
          text: null,
        },
        tickPositions: [0],
      },
      xAxis: {
        title: {
          text: null,
        },
        labels: {
          enabled: false,
        },
        startOnTick: false,
        endOnTick: false,
        lineWidth: 0,
        minorGridLineWidth: 0,
        tickPositions: [],
      },
      credits: {
        enabled: false,
      },
      legend: {
        enabled: false,
      },
      tooltip: {
        formatter: function () {
          return "<b>" + this.key + ": </b>" + this.y.toFixed(2) + "%";
        },
      },
      plotOptions: {
        series: {
          animation: false,
          lineWidth: 1,
          shadow: false,
          states: {
            hover: {
              lineWidth: 1,
            },
          },
          marker: {
            radius: 1,
            states: {
              hover: {
                radius: 2,
              },
            },
          },
        },
        column: {
          color:"#009933",
          negativeColor: "#F26B52",
        },
      },
      series: [
        {
          type: undefined,
          name: "Line 1",
          data: chartData,
        },
      ],
    });
  }

  updatePerfChartData(chartData: any) {
    if (!chartData || !this.fundDetails) {
      return;
    }
    
    var fundName = this.fundDetails.fundName;
    this.performanceChartConfig = new Chart({
      chart: {
        plotBorderWidth: null,
        plotShadow: false,
        type: "line",
        zooming: {
          type: 'x'
        }
      },
      title: {
        text: "",
      },
      yAxis: {
        title: {text: this.translateService.instant('return.percentage'),},
        labels: {
          format: '{value}%'
        }
      },
      xAxis: {
        type: "datetime",
        crosshair: true
      },
      credits: {
        href: "/fsm/main/home",
        text: "",
      },
      legend: {
        enabled: false,
      },
      tooltip: {
        formatter: function () {
          return (
            dateFormat("%d %b %Y", this.x as number) +
            '<br/><span style="color: ' + '">' +
            fundName +
            ": </span><b>" +
            this.y + "%" +
            "</b>"
          );
        },
      },
      plotOptions: {
        area: {
          fillColor: {
            linearGradient: { x1: 0, y1: 0, x2: 0, y2: 1 },
            stops: [
              [0, getOptions().colors[0] as string],
              [1, "#ffffff"],
            ],
          },
        },
        series: {
          marker: {
            enabled: false,
          },
        },
      },
      series: [
        {
          type: undefined,
          name: this.fundDetails.fundName,
          data: chartData,
        },
      ],
    });
  }

  updateFundConsistChartData(chartData: any) {
    this.fundConsistChart = new Chart({
      chart: {
        type: "column",
        zooming: {
          type: 'x'
        },
        backgroundColor: null,
        plotBackgroundColor: null,
        borderWidth: 0,
        width: 150,
        height: 200,
        style: {
          overflow: "visible",
        },
      },
      title: {
        text: "",
      },
      yAxis: {
        endOnTick: false,
        startOnTick: false,
        labels: {
          enabled: false,
        },
        title: {
          text: null,
        },
        tickPositions: [0],
      },
      xAxis: {
        title: {
          text: null,
        },
        labels: {
          enabled: false,
        },
        startOnTick: false,
        endOnTick: false,
        lineWidth: 0,
        minorGridLineWidth: 0,
        tickPositions: [],
      },
      credits: {
        enabled: false,
      },
      legend: {
        enabled: false,
      },
      tooltip: {
        formatter: function () {
          return "<b>" + this.key + ": </b>" + this.y.toFixed(2) + "%";
        },
      },
      plotOptions: {
        series: {
          animation: false,
          lineWidth: 1,
          shadow: false,
          states: {
            hover: {
              lineWidth: 1,
            },
          },
          marker: {
            radius: 1,
            states: {
              hover: {
                radius: 2,
              },
            },
          },
        },
        column: {
          color:"#009933",
          negativeColor: "#F26B52",
        },
      },
      series: [
        {
          type: undefined,
          name: "Line 1",
          data: chartData,
        },
      ],
    });
  }

  getPerfChartConfig() {
    return {
      options: {
        chart: {
          plotBorderWidth: null,
          plotShadow: false,
          type: "area",
          height: 350,
          zooming: {
            type: 'x'
          }
        },
        title: {
          text: "",
        },
        yAxis: {
          title: {
            text: "",
          },
        },
        xAxis: {
          type: "datetime",
        },
        credits: {
          href: "/fsm/main/home",
          text: "",
        },
        legend: {
          enabled: false,
        },
        tooltip: {
          crosshairs: [true],
        },
        plotOptions: {
          series: {
            marker: {
              enabled: false,
            },
          },
        },
      },
      series: [],
    };
  }

  changeFundReturns(selectedCurrencyType: string) {
    if (selectedCurrencyType == "baseCurrency") {
      this.resultCurrencyType = selectedCurrencyType;
      this.selectFundReturnsPerformance =
        this.fundReturnsResult.performanceLocal;
    } else if (selectedCurrencyType == "foreignCurrency") {
      this.resultCurrencyType = selectedCurrencyType;
      this.selectFundReturnsPerformance =
        this.fundReturnsResult.performanceFund;
    }
  }

  scrollToById(idStr: string) {
    var htmlElement: HTMLElement;
    htmlElement = <HTMLElement>document.getElementById(idStr);
    htmlElement.scrollIntoView({ behavior: "smooth"});
  }

  disabledToDate = (current: Date): boolean => {
    return current < this.dateRangeFromDate;
  };

  viewDailyPriceHistoryModal() {
    this.isDailyPriceHistoryModalVisible = true;
  }

  handleDailyPriceHistoryModalCancel() {
    this.isDailyPriceHistoryModalVisible = false;
  }

  sortDailyPriceHistory(sortAttribute): void {
    Object.keys(this.dailyPriceDataSortMap).forEach(key => this.dailyPriceDataSortMap[key] = key === sortAttribute.key ? sortAttribute.value : null);
    this.displayDailyPriceHistoryData = this.utilService.sort(sortAttribute, this.dailyPriceHistory);
  }

  searchDailyPriceHistory() {
    this.displayDailyPriceHistoryData = this.dailyPriceHistory;
    this.displayDailyPriceHistoryData = this.utilService.search(this.searchDailyPriceHistoryText, this.dailyPriceHistory);
  }

  viewRiskRatingModal() {
    this.isRiskRatingModalVisible = true;
  }

  handleRiskRatingModalCancel() {
    this.isRiskRatingModalVisible = false;
  }

  viewHistoricalDividendModal() {
    this.isHistoricalDividendVisible = true;
  }

  handleHistoricalDividendModal() {
    this.isHistoricalDividendVisible = false;
  }

  searchHistoricalDividend() {
    this.displayHistoricalData = this.dividendHistoryList;
    this.displayHistoricalData = this.utilService.search(this.searchHistoricalDividendText, this.dividendHistoryList);
  }

  sortHistoricalDividend(sortAttribute) {
    Object.keys(this.historicalDividendDataSortMap).forEach(key => this.historicalDividendDataSortMap[key] = key === sortAttribute.key ? sortAttribute.value : null);
    this.displayHistoricalData = this.utilService.sort(sortAttribute, this.dividendHistoryList);
  }

  translateDate(inputDateVal: Date, dateFormat?: string) {
    var selectedLocale = this.globalDataStorage.getStorage('locale') || 'en';
    if (selectedLocale == "en") {
        selectedLocale = "en-us";
    } else if (selectedLocale == "ch") {
        selectedLocale = "zh-tw";
    } else if (selectedLocale == "zh") {
        selectedLocale = "zh-cn";
    }
    return this.datePipe.transform(inputDateVal, dateFormat, null, selectedLocale);
  }

  getRiskRatingImageSrc() {
    var selectedLocale =
      this.globalDataStorage.getSessionStorage("locale") || "en";
    this.riskImageLink =
      "/b2b/fsm/images/" + selectedLocale + "/riskRatingOverview.jpg";
    return this.riskImageLink;
  }

  setExportContent() {
    var tabName = "";
    tabName = this.translateService.instant("historical.dividend.allocated");
    var headerList = [];
    var productList = [];
    headerList = [
      [
        this.translateService.instant("ex.date"),
        this.translateService.instant("dividend.rate.per.unit"),
        this.translateService.instant("annualised.dividend.yield"),
      ],
    ];
    this.excelService.excelContent[this.pageName] = [
      {
        worksheet: tabName,
        header: headerList,
        content: productList,
      },
    ];
    this.displayHistoricalData.forEach((dh) => {
      dh.exDateStr = "" + dh.exDate;
      dh.currency = "" + dh.currency;
      dh.rateStr = "" + dh.rate;
      dh.divYield = "" + dh.divYield;
      var row = [
        this.translateDate(dh.exDate, "longDate"),
        dh.currency + " " + dh.rateStr,
        parseFloat(dh.divYield).toFixed(4),
      ];
      productList.push(row);
    });
    var pdfContentList = JSON.parse(JSON.stringify(productList));
    this.pdfService.pdfContentObj[this.pageName].title = [
      { text: tabName, style: "header" },
    ];
    this.pdfService.pdfContentObj[this.pageName].content = [
      {
        style: "table",
        table: {
          headerRows: 1,
          dontBreakRows: true,
          body: this.getPdfBody(headerList[0], pdfContentList),
        },
      },
    ];
  }

  setExportContent2() {
    var title = [];
    var tabName = "";
    if (this.fundDetails != null) {
      tabName =
        this.fundDetails.managerName +
        "\n" +
        this.fundDetails.fundName +
        "\n" +
        this.translateService.instant("daily.prices.for.the.latest.3.months");
      title = [
        this.fundDetails.managerName,
        this.fundDetails.fundName,
        this.translateService.instant("daily.prices.for.the.latest.3.months"),
      ];
    }
    var headerList = [];
    var productList = [];
    headerList = [
      [
        this.translateService.instant("number"),
        this.translateService.instant("nav.price"),
        this.translateService.instant("dealing.date"),
      ],
    ];
    this.excelService.excelContent[this.pageName] = [
      {
        worksheet: tabName,
        titleArray: title,
        header: headerList,
        content: productList,
      },
    ];
    var i = 1;
    this.displayDailyPriceHistoryData.forEach((dph) => {
      dph.bidPriceTest = "" + dph.bidPrice;
      dph.offerPriceTest = "" + dph.offerPrice;
      dph.showDateTest =
        "" + this.datePipe.transform(dph.dailyPricePk.showDate, "mediumDate");

      var row = [
        i++,
        dph.bidPriceTest,
        this.translateDate(dph.showDateTest, "longDate"),
      ];
      productList.push(row);
    });
    var pdfContentList = JSON.parse(JSON.stringify(productList));
    this.pdfService.pdfContentObj[this.pageName].title = [
      { text: tabName, style: "header" },
    ];
    this.pdfService.pdfContentObj[this.pageName].content = [
      {
        style: "table",
        table: {
          headerRows: 1,
          dontBreakRows: true,
          body: this.getPdfBody(headerList[0], pdfContentList),
        },
      },
    ];
  }

  getPdfBody(headerList, dataList) {
    var body = [];
    var headerBody = [];
    headerList.forEach((h) => {
      headerBody.push({
        text: h,
        style: "tableHeader",
      });
    });
    body.push(headerBody);
    dataList.forEach((dataRow) => {
      body.push(dataRow);
    });
    return body;
  }

  onClickPdfAndPrintButton(event: any, a: any) {
    if (a == "historicalPrice") {
      this.setExportContent2();
    } else {
      this.setExportContent();
    }
    this.export(event);
  }

  async export(action: any) {
    if (action == "pdf" || action == "print") {
      var pdfDefinition = await this.pdfService.getPdfDefinition(this.pageName);
      if (pdfDefinition != null) {
        if (action == "pdf") {
          this.pdfService.isPdfLoadingChange.next(true);
          pdfMake.createPdf(pdfDefinition).download(this.pageName, () => {
            this.pdfService.isPdfLoadingChange.next(false);
          });
        }
        if (action == "print") {
          this.pdfService.isPrintLoadingChange.next(true);
          var pdfDoc = pdfMake.createPdf(pdfDefinition);
          pdfDoc.getBuffer((buffer) => {
            this.pdfService.isPrintLoadingChange.next(false);
            pdfDoc.print();
          });
        }
      }
    } else if (action == "excel") {
      var workbook;
      workbook = await this.excelService.getExcelDefinition(this.pageName);
      if (workbook != null) {
        workbook.xlsx.writeBuffer().then((data) => {
          let blob = new Blob([data], { type: EXCEL_TYPE });
          FileSaver.saveAs(blob, this.pageName + EXCEL_EXTENSION);
        });
      }
    }
  }

  async exportToPdf() {
    this.exportLoading = true;
    var content = document.getElementById("inner-content");
    let HTML_Width = content.offsetWidth;
    let HTML_Height = content.offsetHeight;
    let top_left_margin = 15;
    let PDF_Width = HTML_Width + top_left_margin * 2;
    let PDF_Height = PDF_Width * 1.5 + top_left_margin * 2;
    var opt = {
      margin: 1,
      filename: "Fund Factsheet(" + this.fundId + ").pdf",
      image: { type: "jpeg", quality: 0.98 },
      html2canvas: { scale: 2 },
      jsPDF: { unit: "px", format: [PDF_Width, PDF_Height] },
      pagebreak: { before: ".pageBreak", avoid: ["tr", "td", "img"] },
    };
    await html2pdf().from(content).set(opt).save();
    this.exportLoading = false;
  }

  addAlert() {
    const requestObj = {} as any;
    requestObj.paramSedolnumber = this.fundId;
    requestObj.paramCeiling = this.ceilingPrice;
    requestObj.paramFloor = this.floorPrice;
    this.isLoadingPriceAlert = true;
    this.watchlistService.addAlertUt(requestObj).toPromise().then(res => {
        if (!res.success) {
          this.nzModalService.error({
            nzContent: this.translateService.instant(this.errorMsg[res.message] || 'error'),
          });
        }
      }).finally(() => {
        this.showPriceAlert = false;
        this.isLoadingPriceAlert = false;
      });
  }

  formatFrequencyKey(value: string) {
    return value ? value.replace(/[\,]/g, '').replace(/[^a-z0-9A-Z]/g, '.').toLowerCase() : value;
  }

  sort(sort: { key: string; value: any }): void {
    this.fundTopHoldingsList = this.utilService.sort(sort, this.fundTopHoldingsList);
  }
}