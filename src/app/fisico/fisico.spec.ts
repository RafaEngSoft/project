import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Fisico } from './fisico';

describe('Fisico', () => {
  let component: Fisico;
  let fixture: ComponentFixture<Fisico>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Fisico],
    }).compileComponents();

    fixture = TestBed.createComponent(Fisico);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
