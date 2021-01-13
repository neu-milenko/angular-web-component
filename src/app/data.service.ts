import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';

@Injectable({
  providedIn: 'root'
})
export class DataService {
  apiKey = '7527fd852af74dc4bd9238d5878b86b6';
  get(): any{
    return this.httpClient.get(`https://newsapi.org/v2/top-headlines?sources=techcrunch&apiKey=${this.apiKey}`);
  }

  constructor(private httpClient: HttpClient) { }
}
