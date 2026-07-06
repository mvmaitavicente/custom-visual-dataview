/// <reference path="./typings.d.ts" />

"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import DataView = powerbi.DataView;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import DataViewTable = powerbi.DataViewTable;
import DataViewTableRow = powerbi.DataViewTableRow;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import PrimitiveValue = powerbi.PrimitiveValue;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

import { VisualFormattingSettingsModel } from "./settings";

type RiskLevel = "Bajo" | "Medio" | "Alto" | "Muy Alto";

interface ColumnIndexes {
    region: number;
    provincia: number;
    distrito: number;
    codigoLocal: number;
    idSolicitud: number;
    unidadGerencial: number;
    estadoSolicitud: number;
    nivelRiesgo: number;
    montoInversion: number;
    latitud: number;
    longitud: number;
}

interface SourceRecord {
    region: PrimitiveValue;
    provincia: PrimitiveValue;
    distrito: PrimitiveValue;
    codigoLocal: PrimitiveValue;
    idSolicitud: PrimitiveValue;
    unidadGerencial: PrimitiveValue;
    estadoSolicitud: PrimitiveValue;
    nivelRiesgo: PrimitiveValue;
    montoInversion: number;
    latitud: number;
    longitud: number;
}

interface RiskBuckets {
    bajo: number;
    medio: number;
    alto: number;
    muyAlto: number;
}

interface RegionSummary {
    region: PrimitiveValue;
    filas: number;
    colegiosUnicos: number;
    solicitudesUnicas: number;
    montoTotal: number;
    riesgo: RiskBuckets;
    porcentajeCritico: number;
    scoreRiesgo: number;
    clasificacion: RiskLevel;
}

interface UnitSummary {
    unidad: PrimitiveValue;
    filas: number;
    colegiosUnicos: number;
    solicitudesUnicas: number;
    montoTotal: number;
}

interface StateSummary {
    estado: PrimitiveValue;
    colegiosUnicos: number;
    solicitudesUnicas: number;
}

interface RiskSummary {
    nivel: RiskLevel;
    colegiosUnicos: number;
    solicitudesUnicas: number;
    porcentajeSobreTotal: number;
}

interface AnalyticsResult {
    filas: number;
    colegiosUnicos: number;
    solicitudesUnicas: number;
    regionesUnicas: number;
    provinciasUnicas: number;
    distritosUnicos: number;
    unidadesGerencialesUnicas: number;
    montoTotal: number;
    regionSummaries: RegionSummary[];
    unitSummaries: UnitSummary[];
    stateSummaries: StateSummary[];
    riskSummaries: RiskSummary[];
}

interface PerformanceMetrics {
    updateTotalMs: number;
    readDataViewMs: number;
    findColumnsMs: number;
    buildRecordsMs: number;
    accumulateRowsMs: number;
    buildAnalyticsEngineMs: number;
    renderMs: number;
    fetchWaitMs?: number;
}

interface TableDiagnostics {
    dataViewCount: number;
    hasTable: boolean;
    columnIndexes: ColumnIndexes;
    columns: DataViewMetadataColumn[];
    rows: DataViewTableRow[];
    records: SourceRecord[];
    segmentFilas: number;
    accumulatedFilas: number;
    hasMoreData: boolean;
    dataReductionText: string;
    analytics: AnalyticsResult;
    performanceMetrics: PerformanceMetrics;
}

function measure<T>(label: string, fn: () => T): { value: T; ms: number } {
    const start = performance.now();
    const value = fn();
    const end = performance.now();
    console.log(`${label} ms`, end - start);
    return { value, ms: end - start };
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private target: HTMLElement;
    private accumulatedRows: SourceRecord[] = [];
    private accumulatedRowKeys: Set<string> = new Set<string>();
    private lastFetchClickTime?: number;
    private formattingSettings: VisualFormattingSettingsModel = new VisualFormattingSettingsModel();
    private formattingSettingsService: FormattingSettingsService;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.target = options.element;
        this.formattingSettingsService = new FormattingSettingsService();
        this.target.classList.add("data-view-diagnostic");
    }

    public update(options: VisualUpdateOptions): void {
        const updateStart = performance.now();
        const fetchWaitMs = this.lastFetchClickTime
            ? performance.now() - this.lastFetchClickTime
            : undefined;
        this.lastFetchClickTime = undefined;

        const dataViews: DataView[] = options.dataViews || [];
        const dataView = dataViews[0];

        if (dataView) {
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                dataView
            );
        }

        const diagnostics = this.buildDiagnostics(dataViews, fetchWaitMs);
        this.logDiagnostics(diagnostics);

        const renderStart = performance.now();
        this.render(diagnostics);
        diagnostics.performanceMetrics.renderMs = performance.now() - renderStart;
        diagnostics.performanceMetrics.updateTotalMs = performance.now() - updateStart;
        this.updatePerformancePanel(diagnostics);
        console.table(diagnostics.performanceMetrics);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    private buildDiagnostics(dataViews: DataView[], fetchWaitMs?: number): TableDiagnostics {
        const performanceMetrics: PerformanceMetrics = {
            updateTotalMs: 0,
            readDataViewMs: 0,
            findColumnsMs: 0,
            buildRecordsMs: 0,
            accumulateRowsMs: 0,
            buildAnalyticsEngineMs: 0,
            renderMs: 0,
            fetchWaitMs
        };

        const readDataViewResult = measure("readDataView", () => {
            const dataView = dataViews[0];
            const table = dataView?.table;

            return {
                dataView,
                table,
                columns: table?.columns || [],
                rows: table?.rows || [],
                hasMoreData: !!dataView?.metadata?.segment,
                dataReductionText: this.describeDataReduction(table)
            };
        });
        performanceMetrics.readDataViewMs = readDataViewResult.ms;

        const { columns, rows, table, hasMoreData, dataReductionText } = readDataViewResult.value;
        const findColumnsResult = measure("findColumns", () => this.findColumnIndexes(columns));
        performanceMetrics.findColumnsMs = findColumnsResult.ms;

        const columnIndexes = findColumnsResult.value;
        const buildRecordsResult = measure("buildRecords", () => this.buildRecords(rows, columnIndexes));
        performanceMetrics.buildRecordsMs = buildRecordsResult.ms;

        const records = buildRecordsResult.value;
        const accumulateResult = measure("accumulateRows", () => this.accumulateRecords(records));
        performanceMetrics.accumulateRowsMs = accumulateResult.ms;

        const analyticsResult = measure("buildAnalyticsEngine", () => this.buildAnalyticsEngine(this.accumulatedRows));
        performanceMetrics.buildAnalyticsEngineMs = analyticsResult.ms;

        return {
            dataViewCount: dataViews.length,
            hasTable: !!table,
            columnIndexes,
            columns,
            rows,
            records,
            segmentFilas: records.length,
            accumulatedFilas: this.accumulatedRows.length,
            hasMoreData,
            dataReductionText,
            analytics: analyticsResult.value,
            performanceMetrics
        };
    }

    private findColumnIndexes(columns: DataViewMetadataColumn[]): ColumnIndexes {
        return {
            region: this.findColumnIndexByRole(columns, "region"),
            provincia: this.findColumnIndexByRole(columns, "provincia"),
            distrito: this.findColumnIndexByRole(columns, "distrito"),
            codigoLocal: this.findColumnIndexByRole(columns, "codigoLocal"),
            idSolicitud: this.findColumnIndexByRole(columns, "idSolicitud"),
            unidadGerencial: this.findColumnIndexByRole(columns, "unidadGerencial"),
            estadoSolicitud: this.findColumnIndexByRole(columns, "estadoSolicitud"),
            nivelRiesgo: this.findColumnIndexByRole(columns, "nivelRiesgo"),
            montoInversion: this.findColumnIndexByRole(columns, "montoInversion"),
            latitud: this.findColumnIndexByRole(columns, "latitud"),
            longitud: this.findColumnIndexByRole(columns, "longitud")
        };
    }

    private findColumnIndexByRole(columns: DataViewMetadataColumn[], roleName: string): number {
        return columns.findIndex((column: DataViewMetadataColumn) => !!column.roles?.[roleName]);
    }

    private buildRecords(rows: DataViewTableRow[], columnIndexes: ColumnIndexes): SourceRecord[] {
        if (this.hasMissingRole(columnIndexes)) {
            return [];
        }

        return rows.map((row: DataViewTableRow) => ({
            region: row[columnIndexes.region],
            provincia: row[columnIndexes.provincia],
            distrito: row[columnIndexes.distrito],
            codigoLocal: row[columnIndexes.codigoLocal],
            idSolicitud: row[columnIndexes.idSolicitud],
            unidadGerencial: row[columnIndexes.unidadGerencial],
            estadoSolicitud: row[columnIndexes.estadoSolicitud],
            nivelRiesgo: row[columnIndexes.nivelRiesgo],
            montoInversion: this.toNumber(row[columnIndexes.montoInversion]),
            latitud: this.toNumber(row[columnIndexes.latitud]),
            longitud: this.toNumber(row[columnIndexes.longitud])
        }));
    }

    private accumulateRecords(records: SourceRecord[]): void {
        records.forEach((record: SourceRecord) => {
            const recordKey = this.recordKey(record);
            if (this.accumulatedRowKeys.has(recordKey)) {
                return;
            }

            this.accumulatedRowKeys.add(recordKey);
            this.accumulatedRows.push(record);
        });
    }

    private buildAnalyticsEngine(records: SourceRecord[]): AnalyticsResult {
        const colegios = new Set<string>();
        const solicitudes = new Set<string>();
        const regiones = new Set<string>();
        const provincias = new Set<string>();
        const distritos = new Set<string>();
        const unidadesGerenciales = new Set<string>();
        const byRegion = new Map<string, RegionAccumulator>();
        const byUnit = new Map<string, UnitAccumulator>();
        const byState = new Map<string, StateAccumulator>();
        const byRisk = this.createRiskSummaryAccumulators();
        let montoTotal = 0;

        for (const record of records) {
            const colegioKey = this.valueKey(record.codigoLocal);
            const solicitudKey = this.valueKey(record.idSolicitud);
            const regionKey = this.valueKey(record.region);
            const provinciaKey = this.valueKey(record.provincia);
            const distritoKey = this.valueKey(record.distrito);
            const unidadKey = this.valueKey(record.unidadGerencial);
            const estadoKey = this.valueKey(record.estadoSolicitud);
            const riskLevel = this.normalizeRiskLevel(record.nivelRiesgo);

            colegios.add(colegioKey);
            solicitudes.add(solicitudKey);
            regiones.add(regionKey);
            provincias.add(provinciaKey);
            distritos.add(distritoKey);
            unidadesGerenciales.add(unidadKey);
            montoTotal += record.montoInversion;

            this.updateRegionAccumulator(byRegion, regionKey, record, colegioKey, solicitudKey, riskLevel);
            this.updateUnitAccumulator(byUnit, unidadKey, record, colegioKey, solicitudKey);
            this.updateStateAccumulator(byState, estadoKey, record, colegioKey, solicitudKey);
            this.updateRiskAccumulator(byRisk, riskLevel, colegioKey, solicitudKey);
        }

        return {
            filas: records.length,
            colegiosUnicos: colegios.size,
            solicitudesUnicas: solicitudes.size,
            regionesUnicas: regiones.size,
            provinciasUnicas: provincias.size,
            distritosUnicos: distritos.size,
            unidadesGerencialesUnicas: unidadesGerenciales.size,
            montoTotal,
            regionSummaries: this.buildRegionSummaries(byRegion),
            unitSummaries: this.buildUnitSummaries(byUnit),
            stateSummaries: this.buildStateSummaries(byState),
            riskSummaries: this.buildRiskSummaries(byRisk, colegios.size)
        };
    }

    private updateRegionAccumulator(
        byRegion: Map<string, RegionAccumulator>,
        regionKey: string,
        record: SourceRecord,
        colegioKey: string,
        solicitudKey: string,
        riskLevel: RiskLevel
    ): void {
        const item = byRegion.get(regionKey) || {
            region: record.region,
            filas: 0,
            colegios: new Set<string>(),
            solicitudes: new Set<string>(),
            montoTotal: 0,
            riesgo: this.emptyRiskBuckets(),
            weightedRiskTotal: 0
        };

        item.filas += 1;
        item.colegios.add(colegioKey);
        item.solicitudes.add(solicitudKey);
        item.montoTotal += record.montoInversion;
        this.incrementRiskBucket(item.riesgo, riskLevel);
        item.weightedRiskTotal += this.riskScore(riskLevel);
        byRegion.set(regionKey, item);
    }

    private updateUnitAccumulator(
        byUnit: Map<string, UnitAccumulator>,
        unidadKey: string,
        record: SourceRecord,
        colegioKey: string,
        solicitudKey: string
    ): void {
        const item = byUnit.get(unidadKey) || {
            unidad: record.unidadGerencial,
            filas: 0,
            colegios: new Set<string>(),
            solicitudes: new Set<string>(),
            montoTotal: 0
        };

        item.filas += 1;
        item.colegios.add(colegioKey);
        item.solicitudes.add(solicitudKey);
        item.montoTotal += record.montoInversion;
        byUnit.set(unidadKey, item);
    }

    private updateStateAccumulator(
        byState: Map<string, StateAccumulator>,
        estadoKey: string,
        record: SourceRecord,
        colegioKey: string,
        solicitudKey: string
    ): void {
        const item = byState.get(estadoKey) || {
            estado: record.estadoSolicitud,
            colegios: new Set<string>(),
            solicitudes: new Set<string>()
        };

        item.colegios.add(colegioKey);
        item.solicitudes.add(solicitudKey);
        byState.set(estadoKey, item);
    }

    private updateRiskAccumulator(
        byRisk: Map<RiskLevel, RiskAccumulator>,
        riskLevel: RiskLevel,
        colegioKey: string,
        solicitudKey: string
    ): void {
        const item = byRisk.get(riskLevel);
        if (!item) {
            return;
        }

        item.colegios.add(colegioKey);
        item.solicitudes.add(solicitudKey);
    }

    private buildRegionSummaries(byRegion: Map<string, RegionAccumulator>): RegionSummary[] {
        return Array.from(byRegion.values())
            .map((item: RegionAccumulator) => {
                const scoreRiesgo = item.filas ? item.weightedRiskTotal / item.filas : 0;

                return {
                    region: item.region,
                    filas: item.filas,
                    colegiosUnicos: item.colegios.size,
                    solicitudesUnicas: item.solicitudes.size,
                    montoTotal: item.montoTotal,
                    riesgo: item.riesgo,
                    porcentajeCritico: item.filas ? ((item.riesgo.alto + item.riesgo.muyAlto) / item.filas) * 100 : 0,
                    scoreRiesgo,
                    clasificacion: this.classifyRiskScore(scoreRiesgo)
                };
            })
            .sort((left: RegionSummary, right: RegionSummary) => this.displayValue(left.region).localeCompare(this.displayValue(right.region)));
    }

    private buildUnitSummaries(byUnit: Map<string, UnitAccumulator>): UnitSummary[] {
        return Array.from(byUnit.values())
            .map((item: UnitAccumulator) => ({
                unidad: item.unidad,
                filas: item.filas,
                colegiosUnicos: item.colegios.size,
                solicitudesUnicas: item.solicitudes.size,
                montoTotal: item.montoTotal
            }))
            .sort((left: UnitSummary, right: UnitSummary) => this.displayValue(left.unidad).localeCompare(this.displayValue(right.unidad)));
    }

    private buildStateSummaries(byState: Map<string, StateAccumulator>): StateSummary[] {
        return Array.from(byState.values())
            .map((item: StateAccumulator) => ({
                estado: item.estado,
                colegiosUnicos: item.colegios.size,
                solicitudesUnicas: item.solicitudes.size
            }))
            .sort((left: StateSummary, right: StateSummary) => this.displayValue(left.estado).localeCompare(this.displayValue(right.estado)));
    }

    private buildRiskSummaries(byRisk: Map<RiskLevel, RiskAccumulator>, totalColegios: number): RiskSummary[] {
        return Array.from(byRisk.values()).map((item: RiskAccumulator) => ({
            nivel: item.nivel,
            colegiosUnicos: item.colegios.size,
            solicitudesUnicas: item.solicitudes.size,
            porcentajeSobreTotal: totalColegios ? (item.colegios.size / totalColegios) * 100 : 0
        }));
    }

    private logDiagnostics(diagnostics: TableDiagnostics): void {
        console.log("options.dataViews.length", diagnostics.dataViewCount);
        console.log("rows.length", diagnostics.rows.length);
        console.log("segment rows", diagnostics.records.length);
        console.log("accumulated rows", this.accumulatedRows.length);
        console.log("hasMoreData", diagnostics.hasMoreData);
        console.log("columns with roles", diagnostics.columns.map((column: DataViewMetadataColumn, index: number) => ({
            index,
            displayName: column.displayName,
            queryName: column.queryName,
            roles: column.roles
        })));
        console.log("first 10 rows", diagnostics.rows.slice(0, 10));
        console.log("records first 10", diagnostics.records.slice(0, 10));
        console.log("analytics", diagnostics.analytics);
        console.log("data reduction", {
            hasMoreData: diagnostics.hasMoreData,
            detail: diagnostics.dataReductionText
        });
    }

    private render(diagnostics: TableDiagnostics): void {
        this.target.replaceChildren();

        const title = document.createElement("h2");
        title.textContent = "DataView diagnostics";
        this.target.appendChild(title);

        this.appendDataViewPanel(diagnostics);
        this.appendPerformancePanel(diagnostics);

        if (!diagnostics.hasTable) {
            this.appendMessage("No table DataView received.");
            return;
        }

        if (this.hasMissingRole(diagnostics.columnIndexes)) {
            this.appendMessage("Missing one or more required roles for the 80K demo.");
            this.appendColumnsTable(diagnostics.columns);
            return;
        }

        if (diagnostics.hasMoreData) {
            this.appendFetchMoreButton();
        }

        this.appendGlobalTotalsPanel(diagnostics.analytics);
        this.appendRegionSummaryTable(diagnostics.analytics);
        this.appendUnitSummaryTable(diagnostics.analytics.unitSummaries);
        this.appendRiskSummaryTable(diagnostics.analytics.riskSummaries);
        this.appendStateSummaryTable(diagnostics.analytics.stateSummaries);
        this.appendColumnsTable(diagnostics.columns);
    }

    private appendDataViewPanel(diagnostics: TableDiagnostics): void {
        this.appendKeyValueList([
            ["options.dataViews.length", diagnostics.dataViewCount.toString()],
            ["table DataView", diagnostics.hasTable ? "si" : "no"],
            ["filas del segmento actual", diagnostics.segmentFilas.toString()],
            ["filas acumuladas", diagnostics.accumulatedFilas.toString()],
            ["hay mas datos / segment", diagnostics.hasMoreData ? "si" : "no"],
            ["dataReduction", diagnostics.dataReductionText]
        ]);
    }

    private appendGlobalTotalsPanel(analytics: AnalyticsResult): void {
        const title = document.createElement("h2");
        title.textContent = "Totales globales";
        this.target.appendChild(title);

        this.appendKeyValueList([
            ["filas", analytics.filas.toString()],
            ["colegios unicos", analytics.colegiosUnicos.toString()],
            ["solicitudes unicas", analytics.solicitudesUnicas.toString()],
            ["regiones", analytics.regionesUnicas.toString()],
            ["provincias", analytics.provinciasUnicas.toString()],
            ["distritos", analytics.distritosUnicos.toString()],
            ["unidades gerenciales", analytics.unidadesGerencialesUnicas.toString()],
            ["monto total", this.formatNumber(analytics.montoTotal)]
        ]);
    }

    private appendPerformancePanel(diagnostics: TableDiagnostics): void {
        const title = document.createElement("h2");
        title.textContent = "Performance";
        this.target.appendChild(title);

        const metrics = this.getPerformanceDisplayRows(diagnostics);
        const list = document.createElement("dl");
        list.className = "metric-list performance-list";

        metrics.forEach(([key, label, value]: string[]) => {
            const term = document.createElement("dt");
            term.textContent = label;
            list.appendChild(term);

            const description = document.createElement("dd");
            description.dataset.perfKey = key;
            description.textContent = value;
            list.appendChild(description);
        });

        this.target.appendChild(list);
    }

    private updatePerformancePanel(diagnostics: TableDiagnostics): void {
        this.getPerformanceDisplayRows(diagnostics).forEach(([key, , value]: string[]) => {
            const element = this.target.querySelector(`[data-perf-key="${key}"]`);
            if (element) {
                element.textContent = value;
            }
        });
    }

    private getPerformanceDisplayRows(diagnostics: TableDiagnostics): string[][] {
        const metrics = diagnostics.performanceMetrics;
        const processingRate = this.calculateProcessingRate(diagnostics.accumulatedFilas, metrics.updateTotalMs);
        const msPer10k = this.calculateMsPer10k(diagnostics.accumulatedFilas, metrics.updateTotalMs);

        return [
            ["updateTotalMs", "Update total", this.formatMs(metrics.updateTotalMs)],
            ["readDataViewMs", "Read DataView", this.formatMs(metrics.readDataViewMs)],
            ["findColumnsMs", "Find columns", this.formatMs(metrics.findColumnsMs)],
            ["buildRecordsMs", "Build records", this.formatMs(metrics.buildRecordsMs)],
            ["accumulateRowsMs", "Accumulate rows", this.formatMs(metrics.accumulateRowsMs)],
            ["buildAnalyticsEngineMs", "Build analytics engine", this.formatMs(metrics.buildAnalyticsEngineMs)],
            ["renderMs", "Render HTML", this.formatMs(metrics.renderMs)],
            ["fetchWaitMs", "Fetch wait", metrics.fetchWaitMs === undefined ? "n/a" : this.formatMs(metrics.fetchWaitMs)],
            ["segmentFilas", "Rows segment", diagnostics.segmentFilas.toString()],
            ["accumulatedFilas", "Rows accumulated", diagnostics.accumulatedFilas.toString()],
            ["processingRate", "Filas por segundo procesadas", `${processingRate.toFixed(2)} rows/s`],
            ["msPer10k", "Ms por 10,000 filas", this.formatMs(msPer10k)]
        ];
    }

    private appendRegionSummaryTable(analytics: AnalyticsResult): void {
        const title = document.createElement("h2");
        title.textContent = "Resumen por region";
        this.target.appendChild(title);

        const table = this.createTable([
            "Region",
            "Filas",
            "Colegios unicos",
            "Solicitudes unicas",
            "Monto total",
            "Bajo",
            "Medio",
            "Alto",
            "Muy Alto",
            "% critico",
            "Score riesgo",
            "Clasificacion"
        ]);
        const tbody = document.createElement("tbody");

        analytics.regionSummaries.forEach((summary: RegionSummary) => {
            this.appendTableRow(tbody, [
                this.displayValue(summary.region),
                summary.filas.toString(),
                summary.colegiosUnicos.toString(),
                summary.solicitudesUnicas.toString(),
                this.formatNumber(summary.montoTotal),
                summary.riesgo.bajo.toString(),
                summary.riesgo.medio.toString(),
                summary.riesgo.alto.toString(),
                summary.riesgo.muyAlto.toString(),
                this.formatPercent(summary.porcentajeCritico),
                summary.scoreRiesgo.toFixed(2),
                summary.clasificacion
            ]);
        });

        this.appendTableRow(tbody, [
            "TOTAL",
            analytics.filas.toString(),
            analytics.colegiosUnicos.toString(),
            analytics.solicitudesUnicas.toString(),
            this.formatNumber(analytics.montoTotal),
            this.sumRisk("bajo", analytics.regionSummaries).toString(),
            this.sumRisk("medio", analytics.regionSummaries).toString(),
            this.sumRisk("alto", analytics.regionSummaries).toString(),
            this.sumRisk("muyAlto", analytics.regionSummaries).toString(),
            "",
            "",
            ""
        ], "summary-total-row");

        table.appendChild(tbody);
        this.target.appendChild(table);
    }

    private appendUnitSummaryTable(unitSummaries: UnitSummary[]): void {
        const title = document.createElement("h2");
        title.textContent = "Por unidad gerencial";
        this.target.appendChild(title);

        const table = this.createTable(["Unidad", "Filas", "Colegios unicos", "Solicitudes unicas", "Monto total"]);
        const tbody = document.createElement("tbody");

        unitSummaries.forEach((summary: UnitSummary) => {
            this.appendTableRow(tbody, [
                this.displayValue(summary.unidad),
                summary.filas.toString(),
                summary.colegiosUnicos.toString(),
                summary.solicitudesUnicas.toString(),
                this.formatNumber(summary.montoTotal)
            ]);
        });

        table.appendChild(tbody);
        this.target.appendChild(table);
    }

    private appendRiskSummaryTable(riskSummaries: RiskSummary[]): void {
        const title = document.createElement("h2");
        title.textContent = "Por nivel de riesgo";
        this.target.appendChild(title);

        const table = this.createTable(["Nivel", "Colegios unicos", "Solicitudes unicas", "%"]);
        const tbody = document.createElement("tbody");

        riskSummaries.forEach((summary: RiskSummary) => {
            this.appendTableRow(tbody, [
                summary.nivel,
                summary.colegiosUnicos.toString(),
                summary.solicitudesUnicas.toString(),
                this.formatPercent(summary.porcentajeSobreTotal)
            ]);
        });

        table.appendChild(tbody);
        this.target.appendChild(table);
    }

    private appendStateSummaryTable(stateSummaries: StateSummary[]): void {
        const title = document.createElement("h2");
        title.textContent = "Por estado";
        this.target.appendChild(title);

        const table = this.createTable(["Estado", "Colegios unicos", "Solicitudes unicas"]);
        const tbody = document.createElement("tbody");

        stateSummaries.forEach((summary: StateSummary) => {
            this.appendTableRow(tbody, [
                this.displayValue(summary.estado),
                summary.colegiosUnicos.toString(),
                summary.solicitudesUnicas.toString()
            ]);
        });

        table.appendChild(tbody);
        this.target.appendChild(table);
    }

    private appendFetchMoreButton(): void {
        const button = document.createElement("button");
        button.className = "fetch-more-button";
        button.textContent = "Cargar mas datos";
        button.onclick = () => {
            this.lastFetchClickTime = performance.now();
            const accepted = this.host.fetchMoreData();
            console.log("fetchMoreData requested", accepted);
        };

        this.target.appendChild(button);
    }

    private appendColumnsTable(columns: DataViewMetadataColumn[]): void {
        const title = document.createElement("h2");
        title.textContent = "Columnas";
        this.target.appendChild(title);

        const table = this.createTable(["Index", "Column", "Roles"]);
        const tbody = document.createElement("tbody");

        columns.forEach((column: DataViewMetadataColumn, index: number) => {
            this.appendTableRow(tbody, [
                index.toString(),
                column.displayName || "(sin nombre)",
                JSON.stringify(column.roles || {})
            ]);
        });

        table.appendChild(tbody);
        this.target.appendChild(table);
    }

    private appendKeyValueList(metrics: string[][]): void {
        const list = document.createElement("dl");
        list.className = "metric-list";

        metrics.forEach(([label, value]: string[]) => {
            const term = document.createElement("dt");
            term.textContent = label;
            list.appendChild(term);

            const description = document.createElement("dd");
            description.textContent = value;
            list.appendChild(description);
        });

        this.target.appendChild(list);
    }

    private createTable(headers: string[]): HTMLTableElement {
        const table = document.createElement("table");
        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");

        headers.forEach((label: string) => {
            const th = document.createElement("th");
            th.textContent = label;
            headerRow.appendChild(th);
        });

        thead.appendChild(headerRow);
        table.appendChild(thead);

        return table;
    }

    private appendTableRow(tbody: HTMLTableSectionElement, values: string[], className?: string): void {
        const row = document.createElement("tr");
        if (className) {
            row.className = className;
        }

        values.forEach((value: string) => {
            const cell = document.createElement("td");
            cell.textContent = value;
            row.appendChild(cell);
        });

        tbody.appendChild(row);
    }

    private appendMessage(message: string): void {
        const paragraph = document.createElement("p");
        paragraph.textContent = message;
        this.target.appendChild(paragraph);
    }

    private hasMissingRole(columnIndexes: ColumnIndexes): boolean {
        return Object.keys(columnIndexes).some((key: string) => columnIndexes[key as keyof ColumnIndexes] < 0);
    }

    private describeDataReduction(table?: DataViewTable): string {
        if (!table?.columns.length) {
            return "(sin columnas)";
        }

        return `columns=${table.columns.length}, rows=${table.rows?.length || 0}, window=30000`;
    }

    private createRiskSummaryAccumulators(): Map<RiskLevel, RiskAccumulator> {
        return new Map<RiskLevel, RiskAccumulator>([
            ["Bajo", { nivel: "Bajo", colegios: new Set<string>(), solicitudes: new Set<string>() }],
            ["Medio", { nivel: "Medio", colegios: new Set<string>(), solicitudes: new Set<string>() }],
            ["Alto", { nivel: "Alto", colegios: new Set<string>(), solicitudes: new Set<string>() }],
            ["Muy Alto", { nivel: "Muy Alto", colegios: new Set<string>(), solicitudes: new Set<string>() }]
        ]);
    }

    private emptyRiskBuckets(): RiskBuckets {
        return {
            bajo: 0,
            medio: 0,
            alto: 0,
            muyAlto: 0
        };
    }

    private incrementRiskBucket(buckets: RiskBuckets, riskLevel: RiskLevel): void {
        if (riskLevel === "Bajo") {
            buckets.bajo += 1;
        } else if (riskLevel === "Medio") {
            buckets.medio += 1;
        } else if (riskLevel === "Alto") {
            buckets.alto += 1;
        } else {
            buckets.muyAlto += 1;
        }
    }

    private normalizeRiskLevel(value: PrimitiveValue): RiskLevel {
        const normalized = this.displayValue(value)
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .trim()
            .toLowerCase();

        if (normalized === "muy alto" || normalized === "muyalto") {
            return "Muy Alto";
        }

        if (normalized === "alto") {
            return "Alto";
        }

        if (normalized === "medio") {
            return "Medio";
        }

        return "Bajo";
    }

    private riskScore(riskLevel: RiskLevel): number {
        if (riskLevel === "Bajo") {
            return 1;
        }

        if (riskLevel === "Medio") {
            return 2;
        }

        if (riskLevel === "Alto") {
            return 3;
        }

        return 4;
    }

    private classifyRiskScore(score: number): RiskLevel {
        if (score <= 1.75) {
            return "Bajo";
        }

        if (score <= 2.5) {
            return "Medio";
        }

        if (score <= 3.25) {
            return "Alto";
        }

        return "Muy Alto";
    }

    private sumRisk(bucket: keyof RiskBuckets, regionSummaries: RegionSummary[]): number {
        return regionSummaries.reduce((total: number, summary: RegionSummary) => total + summary.riesgo[bucket], 0);
    }

    private displayValue(value: PrimitiveValue): string {
        if (value === null) {
            return "(blank)";
        }

        if (value === undefined) {
            return "(undefined)";
        }

        if (value instanceof Date) {
            return value.toISOString();
        }

        return String(value);
    }

    private valueKey(value: PrimitiveValue): string {
        if (value instanceof Date) {
            return `date:${value.toISOString()}`;
        }

        return `${typeof value}:${String(value)}`;
    }

    private toNumber(value: PrimitiveValue): number {
        if (typeof value === "number") {
            return Number.isFinite(value) ? value : 0;
        }

        const parsed = Number(value);
        return Number.isFinite(parsed) ? parsed : 0;
    }

    private formatNumber(value: number): string {
        return value.toLocaleString(undefined, {
            maximumFractionDigits: 2,
            minimumFractionDigits: 2
        });
    }

    private formatPercent(value: number): string {
        return `${value.toFixed(2)}%`;
    }

    private formatMs(value: number): string {
        if (!Number.isFinite(value)) {
            return "n/a";
        }

        return `${value.toFixed(2)} ms`;
    }

    private calculateProcessingRate(rowCount: number, updateTotalMs: number): number {
        if (!rowCount || !updateTotalMs) {
            return 0;
        }

        return (rowCount / updateTotalMs) * 1000;
    }

    private calculateMsPer10k(rowCount: number, updateTotalMs: number): number {
        if (!rowCount) {
            return 0;
        }

        return (updateTotalMs / rowCount) * 10000;
    }

    private recordKey(record: SourceRecord): string {
        return [
            this.valueKey(record.region),
            this.valueKey(record.provincia),
            this.valueKey(record.distrito),
            this.valueKey(record.codigoLocal),
            this.valueKey(record.idSolicitud),
            this.valueKey(record.unidadGerencial),
            this.valueKey(record.estadoSolicitud),
            this.valueKey(record.nivelRiesgo),
            record.montoInversion.toString(),
            record.latitud.toString(),
            record.longitud.toString()
        ].join("|");
    }
}

interface RegionAccumulator {
    region: PrimitiveValue;
    filas: number;
    colegios: Set<string>;
    solicitudes: Set<string>;
    montoTotal: number;
    riesgo: RiskBuckets;
    weightedRiskTotal: number;
}

interface UnitAccumulator {
    unidad: PrimitiveValue;
    filas: number;
    colegios: Set<string>;
    solicitudes: Set<string>;
    montoTotal: number;
}

interface StateAccumulator {
    estado: PrimitiveValue;
    colegios: Set<string>;
    solicitudes: Set<string>;
}

interface RiskAccumulator {
    nivel: RiskLevel;
    colegios: Set<string>;
    solicitudes: Set<string>;
}
