require([
  'esri/Map',
  'esri/views/MapView',
  'esri/layers/FeatureLayer',
  'esri/Graphic',
  'esri/widgets/Legend',
  'esri/widgets/Expand',
  'esri/geometry/geometryEngine',
  'esri/layers/GraphicsLayer',
  'esri/rest/query',
  'esri/geometry/Polyline',
  'esri/geometry/Polygon'
], function (Map, MapView, FeatureLayer, Graphic, Legend, Expand, geometryEngine, GraphicsLayer, query, Polyline, Polygon) {

    // Add external Air Quality FeatureServer layer
    const airQualityLayer = new FeatureLayer({
      url:
        'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Air_Quality/FeatureServer/0',
      title: 'Air Quality',
      outFields: ['*'],
      renderer: {
        type: 'class-breaks',
        field: 'AQI',
        classBreakInfos: [
          {
            minValue: 0,
            maxValue: 50,
            symbol: {
              type: 'simple-marker',
              size: 12,
              color: '#1a9641',
              outline: { color: 'white', width: 1 }
            },
            label: 'Good (0-50)'
          },
          {
            minValue: 51,
            maxValue: 100,
            symbol: {
              type: 'simple-marker',
              size: 12,
              color: '#ffffbf',
              outline: { color: 'white', width: 1 }
            },
            label: 'Moderate (51-100)'
          },
          {
            minValue: 101,
            maxValue: 200,
            symbol: {
              type: 'simple-marker',
              size: 12,
              color: '#fee090',
              outline: { color: 'white', width: 1 }
            },
            label: 'Poor (101-200)'
          },
          {
            minValue: 201,
            maxValue: 1000,
            symbol: {
              type: 'simple-marker',
              size: 12,
              color: '#d7191c',
              outline: { color: 'white', width: 1 }
            },
            label: 'Very Poor (200+)'
          }
        ]
      },
      popupTemplate: {
        title: 'Air Quality Monitoring',
        content: [
          {
            type: 'fields',
            fieldInfos: [
              { fieldName: 'AQI', label: 'Air Quality Index (AQI)' },
              { fieldName: 'PM2_5', label: 'PM2.5 (µg/m³)' },
              { fieldName: 'PM10', label: 'PM10 (µg/m³)' },
              { fieldName: 'NO2', label: 'NO2 (ppb)' },
              { fieldName: 'O3', label: 'O3 (ppb)' },
              { fieldName: 'timestamp', label: 'Timestamp' }
            ]
          }
        ]
      }
    });

    // Add external Industrial Emissions FeatureServer layer
    const industrialEmissionsLayer = new FeatureLayer({
      url:
        'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Industrial_Emissions/FeatureServer/0',
      title: 'Industrial Emissions',
      outFields: ['*'],
      renderer: {
        type: 'unique-value',
        field: 'compliance_status',
        defaultSymbol: {
          type: 'simple-marker',
          size: 10,
          color: '#cccccc',
          outline: { color: 'white', width: 1 }
        },
        uniqueValueInfos: [
          {
            value: 'Compliant',
            symbol: {
              type: 'simple-marker',
              size: 10,
              color: '#1a9641',
              outline: { color: 'white', width: 1 }
            },
            label: 'Compliant'
          },
          {
            value: 'Non-Compliant',
            symbol: {
              type: 'simple-marker',
              size: 10,
              color: '#d7191c',
              outline: { color: 'white', width: 1 }
            },
            label: 'Non-Compliant'
          }
        ],
        visualVariables: [
          {
            type: 'size',
            field: 'annual_CO2_ton',
            minDataValue: 0,
            maxDataValue: 100000,
            minSize: 8,
            maxSize: 25
          }
        ]
      },
      popupTemplate: {
        title: '{facility_name}',
        content: [
          {
            type: 'fields',
            fieldInfos: [
              { fieldName: 'facility_name', label: 'Facility Name' },
              { fieldName: 'compliance_status', label: 'Compliance Status' },
              { fieldName: 'annual_CO2_ton', label: 'Annual CO2 (tons)' },
              { fieldName: 'industry_sector', label: 'Industry Sector' },
              { fieldName: 'location', label: 'Location' },
              { fieldName: 'last_inspection_date', label: 'Last Inspection Date' },
              { fieldName: 'emissions_limit', label: 'Emissions Limit (tons)' }
            ]
          }
        ]
      }
    });

  const map = new Map({ basemap: 'gray-vector', layers: [airQualityLayer, industrialEmissionsLayer] });

  // Fire Hotspots layer: supports heatmap and point symbology
  const fireHotspotsPointRenderer = {
    type: 'simple',
    symbol: {
      type: 'simple-marker',
      color: 'red',
      outline: { color: 'rgba(255,255,255,0.6)', width: 0.5 }
    },
    visualVariables: [
      {
        type: 'size',
        field: 'brightness',
        minDataValue: 0,
        maxDataValue: 300,
        minSize: 6,
        maxSize: 36
      }
    ]
  };

  const fireHotspotsHeatmapRenderer = {
    type: 'heatmap',
    field: 'brightness',
    colorStops: [
      { ratio: 0, color: 'rgba(0,0,0,0)' },
      { ratio: 0.2, color: 'rgba(255,238,173,0.6)' },
      { ratio: 0.4, color: 'rgba(255,204,92,0.6)' },
      { ratio: 0.6, color: 'rgba(255,140,60,0.6)' },
      { ratio: 0.8, color: 'rgba(255,60,40,0.7)' },
      { ratio: 1, color: 'rgba(180,0,20,0.8)' }
    ],
    blurRadius: 12,
    maxPixelIntensity: 300,
    minPixelIntensity: 0
  };

  const fireHotspotsLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Fire_Hotspots/FeatureServer/0',
    title: 'Fire Hotspots',
    outFields: ['*'],
    visible: false,
    // default to heatmap; switch to point symbology by assigning the point renderer
    renderer: fireHotspotsHeatmapRenderer,
    popupTemplate: {
      title: 'Fire Hotspot',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'brightness', label: 'Brightness' },
            { fieldName: 'confidence', label: 'Confidence' },
            { fieldName: 'frp', label: 'Fire Radiative Power (FRP)' },
            { fieldName: 'acq_date', label: 'Acquisition Date' },
            { fieldName: 'acq_time', label: 'Acquisition Time' }
          ]
        }
      ]
    }
  });

  const floodRiskLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Flood_Risk_Area/FeatureServer/0',
    title: 'Flood Risk',
    outFields: ['*'],
    visible: false,
    renderer: {
      type: 'unique-value',
      field: 'risk_level',
      defaultSymbol: {
        type: 'simple-fill',
        color: 'rgba(200, 200, 200, 0.4)',
        outline: { color: 'rgba(100, 100, 100, 0.8)', width: 1 }
      },
      uniqueValueInfos: [
        {
          value: 'Low',
          symbol: {
            type: 'simple-fill',
            color: 'rgba(144, 238, 144, 0.5)',
            outline: { color: 'rgba(34, 139, 34, 0.8)', width: 1.5 }
          },
          label: 'Low Risk'
        },
        {
          value: 'Medium',
          symbol: {
            type: 'simple-fill',
            color: 'rgba(255, 215, 0, 0.5)',
            outline: { color: 'rgba(184, 134, 11, 0.8)', width: 1.5 }
          },
          label: 'Medium Risk'
        },
        {
          value: 'High',
          symbol: {
            type: 'simple-fill',
            color: 'rgba(255, 140, 0, 0.5)',
            outline: { color: 'rgba(255, 69, 0, 0.8)', width: 1.5 }
          },
          label: 'High Risk'
        },
        {
          value: 'Very High',
          symbol: {
            type: 'simple-fill',
            color: 'rgba(220, 20, 60, 0.6)',
            outline: { color: 'rgba(139, 0, 0, 0.9)', width: 1.5 }
          },
          label: 'Very High Risk'
        }
      ]
    },
    popupTemplate: {
      title: 'Flood Risk Area',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'risk_level', label: 'Risk Level' },
            { fieldName: 'flood_frequency', label: 'Flood Frequency' },
            { fieldName: 'area_name', label: 'Area Name' },
            { fieldName: 'affected_population', label: 'Affected Population' }
          ]
        }
      ]
    }
  });

  const populationDensityLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Population_Density/FeatureServer',
    title: 'Population Density',
    outFields: ['*'],
    visible: false,
    renderer: {
      type: 'class-breaks',
      field: 'density_per_km2',
      classBreakInfos: [
        {
          minValue: 0,
          maxValue: 50,
          symbol: {
            type: 'simple-marker',
            size: 8,
            color: 'rgba(240, 240, 240, 0.8)',
            outline: { color: 'rgba(150, 150, 150, 0.8)', width: 0.5 }
          },
          label: 'Very Low (0-50)'
        },
        {
          minValue: 51,
          maxValue: 150,
          symbol: {
            type: 'simple-marker',
            size: 10,
            color: 'rgba(200, 220, 240, 0.8)',
            outline: { color: 'rgba(100, 150, 200, 0.8)', width: 0.5 }
          },
          label: 'Low (51-150)'
        },
        {
          minValue: 151,
          maxValue: 300,
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(144, 176, 224, 0.8)',
            outline: { color: 'rgba(70, 130, 180, 0.8)', width: 0.5 }
          },
          label: 'Medium (151-300)'
        },
        {
          minValue: 301,
          maxValue: 500,
          symbol: {
            type: 'simple-marker',
            size: 14,
            color: 'rgba(70, 130, 180, 0.8)',
            outline: { color: 'rgba(25, 80, 140, 0.8)', width: 0.5 }
          },
          label: 'High (301-500)'
        },
        {
          minValue: 501,
          maxValue: 10000,
          symbol: {
            type: 'simple-marker',
            size: 16,
            color: 'rgba(25, 80, 140, 1)',
            outline: { color: 'rgba(10, 40, 100, 0.8)', width: 0.5 }
          },
          label: 'Very High (500+)'
        }
      ]
    },
    popupTemplate: {
      title: 'Population Density',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'density_per_km2', label: 'Population Density (per km²)' },
            { fieldName: 'total_population', label: 'Total Population' },
            { fieldName: 'area_name', label: 'Area Name' },
            { fieldName: 'region', label: 'Region' }
          ]
        }
      ]
    }
  });

  const hospitalsLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Hospitals_Clinics/FeatureServer/0',
    title: 'Hospitals & Clinics',
    outFields: ['*'],
    visible: false,
    renderer: {
      type: 'unique-value',
      field: 'type',
      defaultSymbol: {
        type: 'simple-marker',
        size: 10,
        color: 'gray',
        outline: { color: 'white', width: 1 }
      },
      uniqueValueInfos: [
        {
          value: 'Hospital',
          symbol: {
            type: 'text',
            color: '#d7191c',
            text: 'H',
            font: { size: 16, weight: 'bold' },
            haloColor: 'white',
            haloSize: 2
          },
          label: 'Hospital'
        },
        {
          value: 'Clinic',
          symbol: {
            type: 'text',
            color: '#0099ff',
            text: 'C',
            font: { size: 16, weight: 'bold' },
            haloColor: 'white',
            haloSize: 2
          },
          label: 'Clinic'
        }
      ]
    },
    popupTemplate: {
      title: '{name}',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'type', label: 'Type' },
            { fieldName: 'name', label: 'Name' },
            { fieldName: 'bed_capacity', label: 'Bed Capacity' },
            { fieldName: 'emergency_available', label: 'Emergency Available' },
            { fieldName: 'lat', label: 'Latitude' },
            { fieldName: 'lon', label: 'Longitude' }
          ]
        }
      ]
    }
  });

  const esgComplianceLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/ESG_Compliance/FeatureServer/0',
    title: 'ESG Compliance',
    outFields: ['*'],
    visible: false,
    renderer: {
      type: 'class-breaks',
      field: 'overall_score',
      classBreakInfos: [
        {
          minValue: 0,
          maxValue: 59,
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(215, 25, 28, 0.8)',
            outline: { color: 'rgba(165, 0, 38, 0.8)', width: 1 }
          },
          label: '<60'
        },
        {
          minValue: 60,
          maxValue: 79,
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(253, 174, 97, 0.8)',
            outline: { color: 'rgba(215, 69, 0, 0.8)', width: 1 }
          },
          label: '60-79'
        },
        {
          minValue: 80,
          maxValue: 100,
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(26, 150, 65, 0.8)',
            outline: { color: 'rgba(0, 104, 55, 0.8)', width: 1 }
          },
          label: '80-100'
        }
      ]
    },
    popupTemplate: {
      title: 'ESG Compliance',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'overall_score', label: 'Overall Score' },
            { fieldName: 'entity_name', label: 'Entity' },
            { fieldName: 'assessment_date', label: 'Assessment Date' }
          ]
        }
      ]
    }
  });

  const environmentalViolationsLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Environmental_Violations/FeatureServer/0',
    title: 'Environmental Violations',
    outFields: ['*'],
    visible: false,
    renderer: {
      type: 'unique-value',
      field: 'status',
      defaultSymbol: {
        type: 'simple-marker',
        size: 10,
        color: 'rgba(150, 150, 150, 0.8)',
        outline: { color: 'white', width: 1 }
      },
      uniqueValueInfos: [
        {
          value: 'Ongoing',
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(215, 25, 28, 0.8)',
            outline: { color: 'rgba(165, 0, 38, 0.8)', width: 1 }
          },
          label: 'Ongoing'
        },
        {
          value: 'Under Investigation',
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(253, 174, 97, 0.8)',
            outline: { color: 'rgba(215, 69, 0, 0.8)', width: 1 }
          },
          label: 'Under Investigation'
        },
        {
          value: 'Resolved',
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(26, 150, 65, 0.8)',
            outline: { color: 'rgba(0, 104, 55, 0.8)', width: 1 }
          },
          label: 'Resolved'
        }
      ]
    },
    popupTemplate: {
      title: 'Environmental Violation',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'case_id', label: 'Case ID' },
            { fieldName: 'company', label: 'Company' },
            { fieldName: 'violation_type', label: 'Violation Type' },
            { fieldName: 'status', label: 'Status' },
            { fieldName: 'fine_amount', label: 'Fine Amount' },
            { fieldName: 'lat', label: 'Latitude' },
            { fieldName: 'lon', label: 'Longitude' }
          ]
        }
      ]
    }
  });

  map.addMany([fireHotspotsLayer, floodRiskLayer, populationDensityLayer, hospitalsLayer, esgComplianceLayer, environmentalViolationsLayer]);
  airQualityLayer.visible = false;
  industrialEmissionsLayer.visible = false;

  // Graphics layers for buffer and measurement tools (added after feature layers to be on top)
  const bufferGraphicsLayer = new GraphicsLayer();
  map.add(bufferGraphicsLayer);

  const measurementGraphicsLayer = new GraphicsLayer();
  map.add(measurementGraphicsLayer);
  console.log('Graphics layers added to map');

  const view = new MapView({
    container: 'viewDiv',
    map: map,
    // extent that covers peninsular Malaysia and Sabah/Sarawak
    extent: {
      xmin: 99.0,
      ymin: -1.5,
      xmax: 119.0,
      ymax: 7.5,
      spatialReference: { wkid: 4326 }
    }
  });

  view.when(function() {
    console.log('MapView is ready, spatial reference:', view.spatialReference, 'wkid:', view.spatialReference.wkid);
    console.log('Map spatial reference:', map.spatialReference);
  });

  const legend = new Legend({ view: view });
  view.ui.add(new Expand({ view: view, content: legend, expanded: true }), 'top-right');


  // Layer mapping for checkboxes
  const layerMap = {
    'air-quality': airQualityLayer,
    'fire-hotspots': fireHotspotsLayer,
    'industrial-emissions': industrialEmissionsLayer,
    'flood-risk': floodRiskLayer,
    'population-density': populationDensityLayer,
    'hospitals': hospitalsLayer,
    'esg-compliance': esgComplianceLayer,
    'environmental-violations': environmentalViolationsLayer
  };

  function handleCheckboxChange(event) {
    const layerId = event.target.value;
    const isChecked = event.target.checked;
    
    if (layerMap[layerId]) {
      layerMap[layerId].visible = isChecked;
    }
  }

  const layerCheckboxes = document.querySelectorAll('.layer-checkbox');
  layerCheckboxes.forEach(function (checkbox) {
    checkbox.addEventListener('change', handleCheckboxChange);
  });

  // ====== COLLAPSIBLE SECTIONS ======
  const collapsibleHeaders = document.querySelectorAll('.collapsible-header');
  collapsibleHeaders.forEach(function (header) {
    header.addEventListener('click', function () {
      const targetId = this.getAttribute('data-target');
      const content = document.getElementById(targetId);
      const indicator = this.querySelector('.toggle-indicator');

      if (content.classList.contains('collapsed')) {
        // Expand
        content.classList.remove('collapsed');
        indicator.textContent = '▼';
      } else {
        // Collapse
        content.classList.add('collapsed');
        indicator.textContent = '▶';
      }
    });
  });

  // Buffer analysis functionality
  let bufferModeActive = false;
  const bufferDistanceInput = document.getElementById('bufferDistance');
  const activateBufferBtn = document.getElementById('activateBuffer');
  const clearBufferBtn = document.getElementById('clearBuffer');
  const bufferStatus = document.getElementById('bufferStatus');
  const bufferPopup = document.getElementById('bufferPopup');
  const togglePopupBtn = document.getElementById('togglePopup');
  const featuresTableBody = document.getElementById('featuresTableBody');
  const featureDetails = document.getElementById('featureDetails');
  const detailsTitle = document.getElementById('detailsTitle');
  const detailsTableHead = document.getElementById('detailsTableHead');
  const detailsTableBody = document.getElementById('detailsTableBody');
  const closeDetailsBtn = document.getElementById('closeDetails');

  activateBufferBtn.addEventListener('click', function () {
    bufferModeActive = !bufferModeActive;
    if (bufferModeActive) {
      activateBufferBtn.textContent = 'Deactivate Buffer Mode';
      activateBufferBtn.style.background = '#d7191c';
      bufferStatus.textContent = 'Click on map to create buffer';
      clearBufferBtn.style.display = 'block';
      view.cursor = 'crosshair';
    } else {
      activateBufferBtn.textContent = 'Activate Buffer Mode';
      activateBufferBtn.style.background = '#17365d';
      bufferStatus.textContent = '';
      clearBufferBtn.style.display = 'none';
      view.cursor = 'auto';
    }
  });

  clearBufferBtn.addEventListener('click', function () {
    bufferGraphicsLayer.removeAll();
    bufferPopup.classList.add('hidden');
    featuresTableBody.innerHTML = '';
    featureDetails.classList.add('hidden');
    bufferStatus.textContent = 'Buffer cleared. Click on map to create new buffer';
  });

  view.on('click', function (event) {
    if (!bufferModeActive) return;

    const clickPoint = event.mapPoint;
    const bufferDistance = parseFloat(bufferDistanceInput.value) || 5;
    const bufferKm = bufferDistance;

    // Clear previous buffer and results
    bufferGraphicsLayer.removeAll();
    featuresTableBody.innerHTML = '';
    featureDetails.classList.add('hidden');

    // Create buffer geometry (convert km to meters: 1 km = 1000 m)
    const bufferGeometry = geometryEngine.buffer(clickPoint, bufferKm, 'kilometers');

    // Add buffer to graphics layer
    const bufferGraphic = new Graphic({
      geometry: bufferGeometry,
      symbol: {
        type: 'simple-fill',
        color: 'rgba(0, 150, 200, 0.1)',
        outline: {
          color: 'rgba(0, 150, 200, 0.8)',
          width: 2
        }
      }
    });
    bufferGraphicsLayer.add(bufferGraphic);

    // Add center point
    const centerGraphic = new Graphic({
      geometry: clickPoint,
      symbol: {
        type: 'simple-marker',
        color: '#d7191c',
        size: 8,
        outline: { color: 'white', width: 2 }
      }
    });
    bufferGraphicsLayer.add(centerGraphic);

    // Query features in buffer for each visible layer
    const resultsData = [];
    const visibleLayers = Object.keys(layerMap).filter(key => layerMap[key].visible);

    if (visibleLayers.length === 0) {
      bufferStatus.textContent = 'No layers selected. Please select at least one layer.';
      return;
    }

    let queriesCompleted = 0;

    visibleLayers.forEach(function (layerId) {
      const layer = layerMap[layerId];
      const queryParams = {
        geometry: bufferGeometry,
        spatialRelationship: 'intersects',
        outFields: ['*'],
        returnGeometry: false
      };

      layer.queryFeatures(queryParams).then(function (results) {
        const featureCount = results.features.length;
        resultsData.push({
          layId: layerId,
          layerName: layer.title,
          count: featureCount,
          features: results.features
        });

        queriesCompleted++;
        if (queriesCompleted === visibleLayers.length) {
          displayResults(resultsData);
        }
      }).catch(function (err) {
        console.error('Query error for ' + layer.title + ':', err);
        queriesCompleted++;
        if (queriesCompleted === visibleLayers.length) {
          displayResults(resultsData);
        }
      });
    });
  });

  function displayResults(resultsData) {
    featuresTableBody.innerHTML = '';
    let totalFeatures = 0;

    resultsData.forEach(function (result) {
      totalFeatures += result.count;
      const row = document.createElement('tr');
      
      const layerCell = document.createElement('td');
      layerCell.textContent = result.layerName;
      row.appendChild(layerCell);
      
      const countCell = document.createElement('td');
      countCell.textContent = result.count;
      row.appendChild(countCell);
      
      const detailsCell = document.createElement('td');
      if (result.count > 0) {
        const detailsBtn = document.createElement('button');
        detailsBtn.textContent = 'View Details';
        detailsBtn.className = 'details-btn';
        detailsBtn.addEventListener('click', function() {
          showFeatureDetails(result.features, result.layerName);
        });
        detailsCell.appendChild(detailsBtn);
      } else {
        detailsCell.textContent = 'No features';
      }
      row.appendChild(detailsCell);
      
      featuresTableBody.appendChild(row);
    });

    if (totalFeatures > 0) {
      bufferPopup.classList.remove('hidden');
    }
    bufferStatus.textContent = 'Total: ' + totalFeatures + ' features found in ' + parseFloat(bufferDistanceInput.value) + ' km buffer';
  }

  // ====== MEASUREMENT TOOLS ======
  let distanceModeActive = false;
  let areaModeActive = false;
  let distancePoints = [];
  let areaPoints = [];

  const distanceUnitSelect = document.getElementById('distanceUnit');
  const activateDistanceBtn = document.getElementById('activateDistance');
  const clearDistanceBtn = document.getElementById('clearDistance');
  const distanceResult = document.getElementById('distanceResult');

  const areaUnitSelect = document.getElementById('areaUnit');
  const activateAreaBtn = document.getElementById('activateArea');
  const clearAreaBtn = document.getElementById('clearArea');
  const areaResult = document.getElementById('areaResult');

  // Distance Measurement
  activateDistanceBtn.addEventListener('click', function () {
    distanceModeActive = !distanceModeActive;
    if (distanceModeActive) {
      // Deactivate area mode if active
      if (areaModeActive) {
        areaModeActive = false;
        activateAreaBtn.textContent = 'Measure Area';
        activateAreaBtn.style.background = '#17365d';
        clearAreaBtn.style.display = 'none';
      }
      activateDistanceBtn.textContent = 'Stop Measuring Distance';
      activateDistanceBtn.style.background = '#d7191c';
      clearDistanceBtn.style.display = 'inline-block';
      distanceResult.textContent = 'Click on map to add points';
      view.cursor = 'crosshair';
      distancePoints = [];
      measurementGraphicsLayer.removeAll();
    } else {
      activateDistanceBtn.textContent = 'Measure Distance';
      activateDistanceBtn.style.background = '#17365d';
      clearDistanceBtn.style.display = 'none';
      view.cursor = 'auto';
    }
  });

  clearDistanceBtn.addEventListener('click', function () {
    distancePoints = [];
    measurementGraphicsLayer.removeAll();
    distanceResult.textContent = '';
    distanceResult.innerHTML = '';
  });

  // Area Measurement
  activateAreaBtn.addEventListener('click', function () {
    areaModeActive = !areaModeActive;
    if (areaModeActive) {
      // Deactivate distance mode if active
      if (distanceModeActive) {
        distanceModeActive = false;
        activateDistanceBtn.textContent = 'Measure Distance';
        activateDistanceBtn.style.background = '#17365d';
        clearDistanceBtn.style.display = 'none';
      }
      activateAreaBtn.textContent = 'Stop Measuring Area';
      activateAreaBtn.style.background = '#d7191c';
      clearAreaBtn.style.display = 'inline-block';
      areaResult.textContent = 'Click on map to add points. Double-click to finish.';
      view.cursor = 'crosshair';
      areaPoints = [];
      measurementGraphicsLayer.removeAll();
    } else {
      activateAreaBtn.textContent = 'Measure Area';
      activateAreaBtn.style.background = '#17365d';
      clearAreaBtn.style.display = 'none';
      view.cursor = 'auto';
    }
  });

  clearAreaBtn.addEventListener('click', function () {
    areaPoints = [];
    measurementGraphicsLayer.removeAll();
    areaResult.textContent = '';
    areaResult.innerHTML = '';
  });

  // Map click handler for measurements
  let lastClickTime = 0;

  view.on('click', function (event) {
    const currentTime = new Date().getTime();
    const isDoubleClick = currentTime - lastClickTime < 300;
    lastClickTime = currentTime;

    if (distanceModeActive) {
      console.log('Distance mode active, click detected');
      const point = event.mapPoint;
      console.log('Distance point:', point.longitude, point.latitude);
      distancePoints.push(point);

      // Add point graphic
      const pointGraphic = new Graphic({
        geometry: point,
        symbol: {
          type: 'simple-marker',
          color: '#0096c8',
          size: 8,
          outline: { color: 'white', width: 2 }
        }
      });
      measurementGraphicsLayer.add(pointGraphic);

      if (distancePoints.length > 1) {
        // Remove previous line if exists
        const existingLines = measurementGraphicsLayer.graphics.filter(function (g) {
          return g.geometry.type === 'polyline';
        });
        existingLines.forEach(function (g) {
          measurementGraphicsLayer.remove(g);
        });

        // Draw line
        console.log('Creating line with points:', distancePoints);
        
        // Try creating polyline with view's SR
        const polyline = new Polyline({
          paths: [distancePoints.map(p => [p.x, p.y])],
          spatialReference: view.spatialReference
        });
        
        console.log('Created polyline with paths:', polyline.paths, 'SR:', polyline.spatialReference);

        const lineGraphic = new Graphic({
          geometry: polyline,
          symbol: {
            type: 'simple-line',
            color: '#0000ff', // Changed to bright blue
            width: 10, // Increased to 10px
            style: 'solid'
          }
        });
        
        console.log('Created line graphic:', lineGraphic);
        measurementGraphicsLayer.add(lineGraphic);
        console.log('Added distance line graphic, total graphics in layer:', measurementGraphicsLayer.graphics.length);
        
        // Log all graphics in the layer
        measurementGraphicsLayer.graphics.forEach(function(g, i) {
          console.log('Graphic ' + i + ':', g.geometry.type, g);
        });

        // Calculate distance
        let totalDistance = 0;
        for (let i = 0; i < distancePoints.length - 1; i++) {
          totalDistance += geometryEngine.distance(distancePoints[i], distancePoints[i + 1], 'kilometers');
        }

        let displayDistance = totalDistance;
        let unit = distanceUnitSelect.value;

        if (unit === 'miles') {
          displayDistance = totalDistance * 0.621371;
        } else if (unit === 'meters') {
          displayDistance = totalDistance * 1000;
        }

        const unitLabel = unit === 'kilometers' ? 'km' : (unit === 'miles' ? 'mi' : 'm');
        distanceResult.innerHTML = '<strong>Distance: ' + displayDistance.toFixed(2) + ' ' + unitLabel + '</strong>';
      } else {
        distanceResult.textContent = 'Points: 1 - Click to add more points';
      }
    }

    if (areaModeActive) {
      console.log('Area mode active, click detected');
      const point = event.mapPoint;
      console.log('Area point:', point.longitude, point.latitude);

      if (!isDoubleClick) {
        areaPoints.push(point);

        // Add point graphic
        const pointGraphic = new Graphic({
          geometry: point,
          symbol: {
            type: 'simple-marker',
            color: '#00a65d',
            size: 8,
            outline: { color: 'white', width: 2 }
          }
        });
        measurementGraphicsLayer.add(pointGraphic);

        if (areaPoints.length >= 2) {
          // Draw line connecting all points
          const lineCoords = areaPoints.map(function (p) {
            return [p.x, p.y];
          });
          const polyline = new Polyline({
            paths: [lineCoords],
            spatialReference: view.spatialReference
          });

          // Remove previous line if exists
          const existingLines = measurementGraphicsLayer.graphics.filter(function (g) {
            return g.geometry.type === 'polyline';
          });
          existingLines.forEach(function (g) {
            measurementGraphicsLayer.remove(g);
          });

          const lineGraphic = new Graphic({
            geometry: polyline,
            symbol: {
              type: 'simple-line',
              color: '#00ff00', // Changed to bright green
              width: 5, // Increased to 5px
              style: 'solid'
            }
          });
          measurementGraphicsLayer.add(lineGraphic);
        }

        if (areaPoints.length >= 3) {
          // Draw polygon
          const polygonCoords = areaPoints.map(function (p) {
            return [p.x, p.y];
          });
          polygonCoords.push([areaPoints[0].x, areaPoints[0].y]); // Close polygon

          const polygon = new Polygon({
            rings: [polygonCoords],
            spatialReference: view.spatialReference
          });

          // Remove previous polygon if exists
          const existingPolygons = measurementGraphicsLayer.graphics.filter(function (g) {
            return g.geometry.type === 'polygon';
          });
          existingPolygons.forEach(function (g) {
            measurementGraphicsLayer.remove(g);
          });

          const polygonGraphic = new Graphic({
            geometry: polygon,
            symbol: {
              type: 'simple-fill',
              color: 'rgba(0, 165, 93, 0.2)',
              outline: {
                color: '#00a65d',
                width: 2
              }
            }
          });
          measurementGraphicsLayer.add(polygonGraphic);

          // Calculate area
          const areaGeom = new Polygon({
            rings: [polygonCoords],
            spatialReference: view.spatialReference
          });
          let areaValue = geometryEngine.geodesicArea(areaGeom, 'square-kilometers');
          let displayArea = areaValue;
          let unit = areaUnitSelect.value;

          if (unit === 'squareMiles') {
            displayArea = areaValue * 0.386102;
          } else if (unit === 'squareMeters') {
            displayArea = areaValue * 1000000;
          }

          const unitLabel = unit === 'squareKilometers' ? 'km²' : (unit === 'squareMiles' ? 'mi²' : 'm²');
          areaResult.innerHTML = '<strong>Area: ' + displayArea.toFixed(2) + ' ' + unitLabel + '</strong><br>Points: ' + areaPoints.length;
        } else if (areaPoints.length === 2) {
          areaResult.textContent = 'Points: 2 - Click to add more points to create area';
        } else {
          areaResult.textContent = 'Points: ' + areaPoints.length + ' - Click to add more points (need minimum 3)';
        }
      } else if (isDoubleClick && areaPoints.length > 2) {
        // Finish area measurement on double-click
        areaModeActive = false;
        activateAreaBtn.textContent = 'Measure Area';
        activateAreaBtn.style.background = '#17365d';
        clearAreaBtn.style.display = 'none';
        view.cursor = 'auto';
        areaResult.textContent = areaResult.innerHTML + '<br>(Measurement complete - click Clear Area to start new measurement)';
      }
    }
  });

  function showFeatureDetails(features, layerName) {
    if (features.length === 0) return;
    
    // Set title
    detailsTitle.textContent = layerName + ' - Feature Details (' + features.length + ' features)';
    
    // Clear previous content
    detailsTableHead.innerHTML = '';
    detailsTableBody.innerHTML = '';
    
    // Get all unique attribute names
    const allAttributes = new Set();
    features.forEach(feature => {
      Object.keys(feature.attributes).forEach(attr => {
        if (feature.attributes[attr] !== null && feature.attributes[attr] !== undefined) {
          allAttributes.add(attr);
        }
      });
    });
    
    const attributeNames = Array.from(allAttributes);
    
    // Create table header
    const headerRow = document.createElement('tr');
    attributeNames.forEach(attr => {
      const th = document.createElement('th');
      th.textContent = attr;
      headerRow.appendChild(th);
    });
    detailsTableHead.appendChild(headerRow);
    
    // Create table rows for each feature
    features.forEach((feature, index) => {
      const row = document.createElement('tr');
      attributeNames.forEach(attr => {
        const cell = document.createElement('td');
        const value = feature.attributes[attr];
        cell.textContent = value !== null && value !== undefined ? value : '';
        row.appendChild(cell);
      });
      detailsTableBody.appendChild(row);
    });
    
    // Show details section
    featureDetails.classList.remove('hidden');
    
    // Expand popup if collapsed and adjust height
    if (!popupExpanded) {
      togglePopupBtn.click();
    } else {
      bufferPopup.style.maxHeight = '80vh';
    }
  }

  // Toggle popup expand/collapse
  let popupExpanded = true;
  togglePopupBtn.addEventListener('click', function() {
    popupExpanded = !popupExpanded;
    if (popupExpanded) {
      bufferPopup.style.maxHeight = featureDetails.classList.contains('hidden') ? '60vh' : '80vh';
      togglePopupBtn.textContent = '−';
    } else {
      bufferPopup.style.maxHeight = '60px';
      togglePopupBtn.textContent = '+';
    }
  });

  // Close details
  closeDetailsBtn.addEventListener('click', function() {
    featureDetails.classList.add('hidden');
    if (popupExpanded) {
      bufferPopup.style.maxHeight = '60vh';
    }
  });
});
